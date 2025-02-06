"""
This module handles the process of fetching data from an SQL Server database and generating
an manuel report on items that are to be handled manually.
It uses pyodbc for database connectivity and jinja2 for templating the HTML report.
"""

from datetime import datetime, timedelta
import pyodbc
from jinja2 import Template


class ListHandler:
    """
    A class to handle the retrieval and processing of 'Manuel' list items from a database,
    and generate an HTML report based on the retrieved data.
    """
    def __init__(self, rpa_db_connection_string):
        """
        Initialize the ManuelListHandler with a database connection string.

        :param rpa_db_connection_string: The connection string for connecting to the SQL Server database.
        """
        self.connection_string = rpa_db_connection_string

    def fetch_data(self, query: str, params: tuple = ()):
        """
        Fetch data from the SQL database using the provided SQL query and parameters.

        :param query: The SQL query to execute for fetching data.
        :param params: A tuple of parameters to pass to the SQL query.
        :return: A list of dictionaries where each dictionary represents a row of data.
        """
        try:
            conn = pyodbc.connect(self.connection_string)
            cursor = conn.cursor()
            cursor.execute(query, params)
            columns = [column[0] for column in cursor.description]
            results = cursor.fetchall()
            data = [dict(zip(columns, row)) for row in results]
            return data
        except pyodbc.Error as e:
            print("Database error:", e)
            return []
        finally:
            conn.close()

    def get_metadata(self):
        """
        Fetch metadata.

        :return: A list of dictionaries containing the filtered items.
        """
        query = """
            SELECT [os2formWebformId]
            FROM [RPA].[journalizing].[Metadata]
            WHERE destination_system = 'Solteq_Tand'
                    AND isActive = 1
        """
        return self.fetch_data(query,)

    def list_items(self, os2form_webform_id: str):
        """
        Fetches items from the database for the previous week based on the execution date.
        This method should run every Monday and return data for the previous week, from Monday to Sunday.

        :param os2form_webform_id: The ID to filter on.
        :return: A list of dictionaries with the filtered rows.
        """
        today = datetime.now()
        last_monday = today - timedelta(days=today.weekday() + 7)
        last_sunday = last_monday + timedelta(days=6)

        last_monday_str = last_monday.strftime('%Y-%m-%d')
        last_sunday_str = last_sunday.strftime('%Y-%m-%d')

        query = """
            SELECT
                [description],
                [form_id],
                [Status],
                CAST([Indsendt dato] AS DATETIME) AS [Indsendt dato],
                [CPR],
                [Navn],
                [Klinik],
                [Adresse],
                [Samletaccept],
                [Journalaccept],
                CAST([last_time_modified] AS DATETIME) AS [last_time_modified]
            FROM
                [RPA].[journalizing].[view_Tandplejen_SolteqTand]
            WHERE
                os2formWebformId = ?
                AND CAST([Indsendt dato] AS DATE) BETWEEN ? AND ?
            ORDER BY
                CAST([Indsendt dato] AS DATETIME) DESC
        """

        parameters = (os2form_webform_id, last_monday_str, last_sunday_str)
        data = self.fetch_data(query, parameters)

        return data

    def filter_empty_columns(self, data):
        """
        Filters out columns from the data (a list of dictionaries) where all values are empty or None.

        :param data: List of dictionaries representing rows from the database.
        :return: List of dictionaries with the "empty" columns removed.
        """
        if not data:
            return data

        keys = list(data[0].keys())

        keys_to_remove = []
        for key in keys:
            if all(not row.get(key) for row in data):
                keys_to_remove.append(key)

        filtered_data = []
        for row in data:
            filtered_row = {k: v for k, v in row.items() if k not in keys_to_remove}
            filtered_data.append(filtered_row)

        return filtered_data

    def generate_list(self):
        """
        Generates a consolidated HTML report containing a table for each os2formWebformId,
        but only if there is actual data for the given ID. If no data exists, it is skipped.

        :return: A string containing the full HTML content, ready to be sent in an email.
        """
        metadata = self.get_metadata()

        all_tables = ""

        for metadata_item in metadata:
            current_id = metadata_item['os2formWebformId']
            print(f"Behandler os2formWebformId: {current_id}")

            list_items = self.list_items(current_id)
            if not list_items:
                continue

            list_items = self.filter_empty_columns(list_items)

            description_header = list_items[0].get('description', f"Webform ID: {current_id}")

            list_items_table = self.convert_to_html_table(list_items)

            all_tables += f"<h3>{description_header}</h3>"
            all_tables += list_items_table
            all_tables += "<br>"

        html_template = """
        <html>
        <head>
            <style>
                table, th, td {
                    border: 1px solid black;
                    border-collapse: collapse;
                    padding: 8px;
                    font-size: 12px;
                }
                h3 {
                    font-size: 16px;
                }
                body {
                    font-size: 12px;
                }
            </style>
        </head>
        <body>
            {{ all_tables | safe }}
        </body>
        </html>
        """

        template = Template(html_template)
        html_content = template.render(all_tables=all_tables)

        return html_content

    def convert_to_html_table(self, data):
        """
        Convert a list of dictionaries into an HTML table format.

        :param data: A list of dictionaries where each dictionary represents a row of data.
        :return: A string representing the data as an HTML table.
        """
        if not data:
            return "<p>Fandt ingen rækker. Ingen data tilgængelig.</p>"

        headers = data[0].keys()
        header_html = ''.join(f'<th>{header}</th>' for header in headers)

        rows_html = ''.join(
            '<tr>' + ''.join(f'<td>{value}</td>' for value in row.values()) + '</tr>'
            for row in data
        )

        html_table = f'<table><thead><tr>{header_html}</tr></thead><tbody>{rows_html}</tbody></table>'
        return html_table

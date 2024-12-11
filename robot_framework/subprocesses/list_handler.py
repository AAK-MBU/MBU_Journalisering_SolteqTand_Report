"""
This module handles the process of fetching data from an SQL Server database and generating
an manuel report on items that are to be handled manually.
It uses pyodbc for database connectivity and jinja2 for templating the HTML report.
"""
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

    def fetch_data(self, query):
        """
        Fetch data from the SQL database using the provided SQL query.

        :param query: The SQL query to execute for fetching data.
        :return: A list of dictionaries where each dictionary represents a row of data.
        """
        conn = pyodbc.connect(self.connection_string)
        cursor = conn.cursor()
        cursor.execute(query)
        columns = [column[0] for column in cursor.description]
        results = cursor.fetchall()
        conn.close()
        data = [dict(zip(columns, row)) for row in results]
        return data

    def list_items(self, status: str):
        """
        Fetch all the items from the database that have the relevent status and the current date.

        :return: A list of dictionaries containing the items for the 'Manuel' or 'Successful' list.
        """
        query = f"""
            SELECT [Formular]
                ,[reference]
                ,[Status]
                ,[Indsendt dato]
                ,[CPR MitId]
                ,[CPR Barn]
                ,[Navn]
                ,[Tandlæge/Tandklinik]
                ,[Adresse]
                ,[Samlet accept]
                ,[Tilladelse til at sende journal]
            FROM [RPA].[rpa].[SolteqTand_tandplejetilbud_privat_klinik]
            WHERE Status = '{status}'
            AND FORMAT(CAST(creation_time AS DATETIME), 'dd-MM-yyyy') = FORMAT(CAST(GETDATE() AS DATETIME), 'dd-MM-yyyy')
            ORDER BY Formular ASC
        """
        return self.fetch_data(query)

    def generate_list(self):
        """
        Generate an HTML report for either 'Manuel' or 'Successful' status items.

        The report includes an HTML table of items retrieved from the database, formatted
        as per the specified list.

        :param list_type: A string indicating the type of list to generate.
                        "manuel" for manuel_list_items(), "successful" for successful_list_items().
        :return: A string representing the generated HTML content.
        """
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
            <h3>Succesfuld</h3>
            {{ list_items_table_ok | safe }}
            <br/><br/>
            <h3>Manuel</h3>
            {{ list_items_table_manuel | safe }}
        </body>
        </html>
        """

        list_items_manuel = self.list_items('Manuel')
        list_items_ok = self.list_items('Successful')
        list_items_table_manuel = self.convert_to_html_table(list_items_manuel)
        list_items_table_ok = self.convert_to_html_table(list_items_ok)

        template = Template(html_template)
        html_content = template.render(
            list_items_table_manuel=list_items_table_manuel,
            list_items_table_ok=list_items_table_ok
        )

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

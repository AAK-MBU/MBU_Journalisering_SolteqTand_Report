"""
This module handles the process of fetching data from an SQL Server database and generating 
an manuel report on items that are to be handled manually.
It uses pyodbc for database connectivity and jinja2 for templating the HTML report.
"""
import pyodbc
from jinja2 import Template


class ManuelListHandler:
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

    def manuel_list_items(self):
        """
        Fetch all the items from the database that have the status 'Manuel' and the current date.

        The query retrieves specific fields such as reference, date, and other form details,
        returning them in a list of dictionaries.

        :return: A list of dictionaries containing the items for the 'Manuel' list.
        """
        query = """
        SELECT	'Tandplejetilbud privat klinik 0-17 år' AS [Formular]
                ,[reference]
                ,FORMAT(CAST(JSON_VALUE(data, '$.entity.completed[0].value') AS DATETIMEOFFSET), 'dd-MM-yyyy') AS [Indsendt dato]
                ,JSON_VALUE(data, '$.data.cpr_nummer') AS [CPR MitId]
                ,JSON_VALUE(data, '$.data.cpr_nummr_barnet_manuelt') AS [CPR]
                ,JSON_VALUE(data, '$.data.navn') AS [Navn]
                ,JSON_VALUE(data, '$.data.tandklinik') AS [Tandklinik]
                ,JSON_VALUE(data, '$.data.jeg_er_indforstaaet_med_reglerne_for_frit_valg_og_herunder_regle') AS [Indforstået med reglerne for frit valg]
                ,JSON_VALUE(data, '$.data.jeg_giver_tilladelse_til_at_tandplejen_aarhus_maa_sende_journal_') AS [Tilladelse til at sende journal]
                ,JSON_VALUE(data, '$.data.jeg_har_laest_og_forstaaet_reglerne_og_oensker_at_mit_barn_0_15') AS [Har læst og forstået reglerne]
        FROM	[RPA].[rpa].[Hub_SolteqTand_tandplejetilbud_privat_klinik_0_17_aar]
        WHERE	process_status = 'Manuel'
        UNION ALL
        SELECT	'Tandplejetilbud privat klinik 18-21 år'
                ,[reference]
                ,FORMAT(CAST(JSON_VALUE(data, '$.entity.completed[0].value') AS DATETIMEOFFSET), 'dd-MM-yyyy') AS [Indsendt dato]
                ,JSON_VALUE(data, '$.data.cpr_nummer') AS [CPR MitId]
                ,JSON_VALUE(data, '$.data.cpr_nummr_barnet_manuelt') AS [CPR]
                ,JSON_VALUE(data, '$.data.navn') AS [Navn]
                ,JSON_VALUE(data, '$.data.tandklinik') AS [Tandklinik]
                ,JSON_VALUE(data, '$.data.jeg_er_indforstaaet_med_reglerne_for_frit_valg_og_herunder_regle') AS [Indforstået med reglerne for frit valg]
                ,JSON_VALUE(data, '$.data.jeg_giver_tilladelse_til_at_tandplejen_aarhus_maa_sende_journal_') AS [Tilladelse til at sende journal]
                ,JSON_VALUE(data, '$.data.jeg_har_laest_og_forstaaet_reglerne_og_oensker_at_mit_barn_0_15') AS [Har læst og forstået reglerne]
        FROM	[RPA].[rpa].[Hub_SolteqTand_tandplejetilbud_privat_klinik_18_21_aar]
        WHERE	process_status = 'Manuel'
        """
        return self.fetch_data(query)

    def generate_list(self):
        """
        Generate an HTML report for all process data, including 'Manuel' status items.

        The report includes an HTML table of items retrieved from the database, formatted
        as per the 'Manuel' list.

        :return: A string representing the generated HTML content.
        """
        manuel_list_items = self.manuel_list_items()

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
            <h3>Manuel liste</h3>
            {{ manuel_list_items_table | safe }}
        </body>
        </html>
        """

        manuel_list_items_table = self.convert_to_html_table(manuel_list_items)

        template = Template(html_template)
        html_content = template.render(
            manuel_list_items_table=manuel_list_items_table,
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

import pdfplumber
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font

class PdfToExcelConverter:
    @staticmethod
    def extract_tables_from_pdf(pdf_path):
        """
        Extrae las tablas de un archivo PDF.

        Args:
            pdf_path (str): Ruta del archivo PDF.

        Returns:
            list: Lista de tablas extraídas.
        """
        tables = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                tables.extend(page.extract_tables())
        return tables

    @staticmethod
    def combine_tables(tables):
        """
        Combina una lista de tablas en un solo DataFrame.

        Args:
            tables (list): Lista de DataFrames que representan las tablas.

        Returns:
            pandas.DataFrame: DataFrame combinado.
        """
        combined_table = pd.concat([pd.DataFrame(table) for table in tables], ignore_index=True)
        return combined_table

    @staticmethod
    def split_columns_by_newline(table):
        """
        Divide las columnas de un DataFrame basado en el separador '\n' en cada celda.

        Args:
            table (pandas.DataFrame): DataFrame que contiene las columnas a dividir.

        Returns:
            pandas.DataFrame: DataFrame con las columnas divididas.
        """
        new_columns = {}
        for column in table.columns:
            new_column = table[column].str.split("\n", expand=True)
            num_new_columns = new_column.shape[1]
            new_column_names = [f"{column}_{i}" for i in range(num_new_columns)]
            new_column.columns = new_column_names
            new_columns[column] = new_column
        combined_table = pd.concat(new_columns.values(), axis=1)
        return combined_table

    @staticmethod
    def remove_rows_starting_with_puesto(table):
        """
        Elimina todas las filas que comienzan con "PUESTO" excepto la primera.

        Args:
            table (pandas.DataFrame): DataFrame que contiene las filas a filtrar.

        Returns:
            pandas.DataFrame: DataFrame con las filas filtradas.
        """
        first_puesto_index = table.index[table.iloc[:, 0].str.startswith("PUESTO")].min()
        table = pd.concat([table.iloc[[first_puesto_index]], table.iloc[first_puesto_index+1:].loc[~table.iloc[first_puesto_index+1:, 0].str.startswith("PUESTO")]], ignore_index=True)
        return table

    @staticmethod
    def apply_bold_font_to_first_row(excel_file):
        """
        Aplica negrita a la primera fila de un archivo Excel.

        Args:
            excel_file (str): Ruta del archivo Excel.

        Returns:
            None
        """
        workbook = load_workbook(excel_file)
        sheet = workbook.active
        for cell in sheet["1"]:
            cell.font = Font(bold=True)
        workbook.save(excel_file)

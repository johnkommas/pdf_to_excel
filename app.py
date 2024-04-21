#  Copyright (c) Ioannis E. Kommas. All Rights Reserved 2024.

import os
import subprocess
import sys
import logging
import pandas as pd
import tabula
from PyPDF2 import PdfReader
import openpyxl

# CONSTANTS
OUTPUT_FILE_NAME = 'output.xlsx'
INPUT_FILE_NAME = 'INVOICE'
cwd = os.getcwd()
FOLDER = f'{cwd}/{INPUT_FILE_NAME}'

"""
    '(top, left, bottom, right)'
"""
area_to_check = {'WRONG AREA': (225, 51, 240, 94),
                 'BAZAAR': (52.3, 420, 63, 466), }

suppliers = {'094384144': {'area': (290, 0, 585, 595),
                           'columns': [1, 0, 3, 6, 7, 8, 9, 10],
                           'names': {1: 'ΠΕΡΙΓΡΑΦΗ',
                                     0: 'ΚΩΔΙΚΟΣ',
                                     3: 'ΠΟΣΟΤΗΤΑ',
                                     6: 'ΤΙΜΗ',
                                     7: 'ΕΚΠΤΩΣΗ Α',
                                     8: 'ΕΚΠΤΩΣΗ Β',
                                     9: 'ΕΦΚ',
                                     10: 'ΚΟΣΤΟΣ'
                                     }
                           },
             }


def setup_pandas_display_options():
    """
    This function sets up the display options for pandas DataFrame output.
    It adjusts the maximum number of displayed columns, the width,
    and the maximum number of displayed rows.
    """
    pd.set_option("display.max_columns", 500)
    pd.set_option("display.width", 1000)
    pd.set_option("display.max_rows", 1000)


def setup_logging():
    """
    This function sets up basic logging configuration which writes logs of
    all severity levels equal to or above INFO to a file named 'app.log'.
    The log format includes logger name, severity level, and the message.
    """
    logging.basicConfig(filename='app.log',
                        filemode='w',
                        format='%(name)s - %(levelname)s - %(message)s',
                        level=logging.INFO)
    logger = logging.getLogger(__name__)
    return logger


def get_files_in_directory(directory_path):
    """
        This function lists all the files in a directory provided
        and returns the first file if any found.

        :param directory_path: str, The path of the directory
        :returns: str or None, The first file in the directory or None for empty directory or directories without files
    """
    files = [f for f in os.listdir(directory_path) if os.path.isfile(os.path.join(directory_path, f))]
    if files:
        print(files[0])
        return files[0]
    else:
        return None


def get_pdf_size(file_path):
    """
        Retrieves the dimensions (width and height) of the first page of a PDF file.

        Parameters:
        file_path (string): A string with the system path to the PDF file.

        Returns:
        tuple: A tuple containing the width and height of the first page. If there is an error, it will raise an Exception.


        """
    logger.info(f"Getting PDF size for {file_path}")
    with open(file_path, 'rb') as f:
        try:
            pdf = PdfReader(f)
            page = pdf.pages[0]
            logger.info(f"Successfully got PDF size for {file_path}")
            return page.mediabox[2], page.mediabox[3]
        except Exception as e:
            logger.error(f"Error occurred while getting PDF size: {e}")
            raise


def adjust_column_width(filename):
    """
        Adjusts the width of columns in an Excel file for better readability. Each column width
        is set to the length of the longest cell value in that column.

        Parameters:
        filename (str): A string containing the system path to the Excel file.

        Returns:
        None. The function modifies the Excel file directly.

        Raises:
        Exception: If there's an error in processing the Excel file, an Exception is raised with a
        description of the error.
    """
    logger.info(f"Adjusting column width for {filename}")

    try:
        wb = openpyxl.load_workbook(filename)
        sheet = wb.active

        for column in sheet.columns:
            max_length = 0
            column = [cell for cell in column]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except Exception as e:
                    pass
            adjusted_width = (max_length + 2)
            sheet.column_dimensions[column[0].column_letter].width = adjusted_width

        wb.save(filename)
        logger.info(f"Successfully adjusted column width for {filename}")
    except Exception as e:
        logger.error(f"Error occurred while adjusting column width: {e}")
        raise


def print_file_size(unique_file):
    """
    Prints the width and height of the first page of a given PDF file to the logger.

    Parameters:
    file (str): A string containing the system path to the PDF file.

    Returns:
    None. The function directly logs the width and height of the PDF file.

    Raises:
    Exception: If there's an error in retrieving the PDF file size, an Exception is raised with a
    description of the error.
    """

    try:
        width, height = get_pdf_size(unique_file)
        logger.info(f'The PDF size is {width} points wide and {height} points high.')
    except Exception as e:
        logger.error(f"Error occurred while printing file size: {e}")
        raise


def create_dataframe(unique_file, area, columns, names):
    """
    Creates a pandas DataFrame from the contents of a given PDF file.

    It reads data from all pages of the PDF, cleans up and rearranges columns,
    and finally concatenates into one unified DataFrame.

    Parameters:
    file (str): A string containing the system path to the PDF file.

    Returns:
    DataFrame: A pandas DataFrame containing data processed from the PDF file.
    """
    CODE_COLUMN = 'ΚΩΔΙΚΟΣ'
    # Create empty list to hold dataframes
    dfs = []

    # Read data from all pages
    df_list = tabula.read_pdf(unique_file, pages='all', area=area, guess=False, pandas_options={'header': None})
    # Iterate over list of dataframes and add each to dfs
    for df in df_list:
        df[3] = df[3].replace({r'[^\d\.\,]': ''}, regex=True)
        df = df[columns]
        dfs.append(df)

    # Concatenate dataframes along rows
    final_df = pd.concat(dfs).reset_index(drop=True)
    final_df = final_df.rename(columns=names)
    logger.info(final_df)
    final_df[CODE_COLUMN] = final_df[CODE_COLUMN].astype(str)
    return final_df


def save_and_open(df):
    """
    Saves a provided pandas DataFrame to an Excel file, adjusts the column widths for readability,
    and opens the file automatically.

    Parameters:
    df (DataFrame): A pandas DataFrame that needs to be saved to an Excel file.

    Returns:
    None. The function saves the DataFrame to an Excel file and opens it.
    """

    # Save your DataFrame to an Excel file
    df.to_excel(OUTPUT_FILE_NAME, index=False)
    # Adjust the column width of the Excel file
    adjust_column_width(OUTPUT_FILE_NAME)
    # Open the Excel file
    subprocess.run(['open', OUTPUT_FILE_NAME], check=True)


def main(unique_supplier_id):
    """
    Launches the main process for the script.

    It gets the current working directory and constructs a file path for 'invoice.pdf' file in it.
    If the file exists, it gets the file size, converts the data into a pandas DataFrame,
    saves the DataFrame into 'output.xlsx', and opens the Excel file.

    If at any point an error occurs, it logs the error message and exits with status code 1.
    """
    area = suppliers[unique_supplier_id].get('area')
    columns = suppliers[unique_supplier_id].get('columns')
    names = suppliers[unique_supplier_id].get('names')

    if not os.path.exists(file):
        logger.error(f"File {file} does not exist.")
        sys.exit(1)
    try:
        print_file_size(file)
        df = create_dataframe(file, area, columns, names)
        save_and_open(df)
    except Exception as e:
        logger.error(f"An error occurred: {e}")
        sys.exit(1)


def read_supplier_unique_number(unique_file, area):
    """
    Retrieves the supplier's unique number from a given area of the PDF file.

    Parameters:
    unique_file (str): A string containing the system path to the PDF file.
    area (tuple): A tuple containing the areas of the page to be read as per the tabula.read_pdf() syntax.

    Returns:
    str: The supplier's unique number as a string. It directly matches a key in the global 'suppliers' dictionary.
         If the supplier's unique number is not found or a pdf reading error occurs, it returns None.

    Note:
    This function uses the tabula.read_pdf() method to read a given area of the PDF file as a dataframe
    The supplier's unique number is assumed to be in the first cell (0,0) of the dataframe.

    """
    try:
        df = tabula.read_pdf(unique_file, pages='all', area=area, guess=False,
                             pandas_options={'header': None, 'dtype': str})[0]
    except Exception as e:
        print('No Dataframe Available Here', e)
        return
    try:
        if not df.empty:
            answer = df[0].values[0]
            if answer in suppliers.keys():
                return answer
            else:
                print("\033[91m {}\033[00m".format(f'{answer} not in keys'))
    except Exception as e:
        print("\033[91m {}\033[00m".format(e))
        return None


def find_suppliers(unique_file):
    """
    Iterates over a predefined list of areas in the PDF file and attempts to find a supplier's unique number in each one.

    Parameters:
    unique_file (str): A string specifying the system path to the PDF file.

    Returns:
    str or None : The supplier's unique number as a string if found. If no supplier number is found in any of the defined areas, will return None.

    Note:
    This function relies on the 'area_to_check' dictionary (defined globally). The keys of this dictionary represent identifiers for the positions in the document, and the values are coordinates for the areas to check on the PDF file.
    These areas are used to extract information which is then tested for presence in the 'suppliers' dictionary.
    The function outputs progress to the console and returns as soon as a valid supplier number is identified.
    """
    for key in area_to_check:
        value = area_to_check[key]
        print(f"Checking: {key} Position")
        unique_supplier_id = read_supplier_unique_number(unique_file, value)
        if unique_supplier_id is not None:
            print("\033[92m {}\033[00m".format(f'Position for {key} Worked!'))
            return unique_supplier_id
        else:
            print(f"Nothing Found Checking Next Key")


if __name__ == "__main__":
    logger = setup_logging()
    setup_pandas_display_options()
    file = f'{FOLDER}/{get_files_in_directory(FOLDER)}'
    supplier_id = find_suppliers(file)
    main(supplier_id)

from openpyxl import load_workbook
import pandas as pd
from pathlib import Path
from typing import List, Tuple
from collections import defaultdict
import os
import re

source_folder = r'C:\Users\Somers\Desktop\fake_patient'
excel_file = 'test.xlsx'
destination_folder = r'C:\Users\Somers\Desktop\test'


divider = os.path.sep if os.path.sep != "\\" else os.path.sep * 2
patient_pattern = ".*" + divider +"(.*?)_(.*?)_(\d{4}?)(\d{2}?)(\d{2}?)_(\d{7}-\d{1}?)_(\d{4}?)(\d{2}?)(\d{2}?)_\d{6}(\d{2}?)(\d{2}?).*"
EXCEL_HEADERS = ["GRK Nummer", "Name", "Geburtsdatum", "OP-Datum", "Patient-ID"]


def get_df_from_excel(filename) -> pd.DataFrame:
    """
    Load the grk patient data from an excel sheet
    Parameters
    ----------
    filename : excelsheet to load

    Returns
    -------
        A pandas dataframe of the data in the excel sheet under the tab "Patienten"
    """
    workbook = load_workbook(filename=filename)
    sheet = workbook['Patienten']
    data = sheet.values
    columns = next(data)[0:]
    workbook.close()
    return pd.DataFrame(data, columns=columns)


def get_patients_from_folders(folder) -> List[Tuple[str, Tuple[str]]]:
    """
    Given a folder, will search for all subfolders with the Storz folder naming and return them.
    Parameters
    ----------
    folder : folder to recursively search through

    Returns
    -------
        A list of tuples each with the folder path of the patient and a tuple of the extracted patient info parsed by regex

    """
    files_folders = [str(x) for x in Path(folder).rglob("*[0-9].*")]
    regex_term = re.compile(patient_pattern)
    directories = set([x.string[:x.regs[-1][1]] for x in map(regex_term.fullmatch, files_folders) if x])
    matches = [x[0] for x in map(regex_term.findall, directories) if x]
    return list(zip(directories, matches))


def convert_patient_info_to_df(patients):
    """ get data as dictionary """
    df_dict = defaultdict(list)
    for patient in patients:
        data = patient[1]
        df_dict['GRK Nummer'].append(None)
        df_dict['Name'].append(f"{data[0]}, {data[1]}")
        df_dict['Geburtsdatum'].append(f"{data[4]}.{data[3]}.{data[2]}")
        df_dict['OP-Datum'].append(f"{data[8]}.{data[7]}.{data[6]}")
        df_dict['Patient-ID'].append(f"{data[5]}")
    return df_dict


def write_dataframe_to_excel(dataframe: pd.DataFrame, excel_file: str):
    """
    writes a pandas dataframe to an excel workbook to the sheet titled "Patienten"
    :param dataframe:
    :param excel_file:
    """
    dataframe.to_excel(excel_writer=excel_file,
                       sheet_name="Patienten",
                       index=False,
                       engine='openpyxl')

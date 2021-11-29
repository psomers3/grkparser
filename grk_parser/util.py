from openpyxl import load_workbook
import pandas as pd
from pathlib import Path
from typing import List, Tuple
from collections import defaultdict
import os
import re

divider = os.path.sep if os.path.sep != "\\" else os.path.sep * 2
patient_patterns = {'full_info': ".*" + divider +"(.*?)_(.*?)_(\d{4}?)(\d{2}?)(\d{2}?)_(.*?)_(\d{4}?)(\d{2}?)(\d{2}?)_\d{6}(\d{2}?)(\d{2}?).*",
                    'name_id_opdate_time': ".*" + divider +"(.*?)_(.*?)_(.*?)_(\d{4}?)(\d{2}?)(\d{2}?)_(\d{2}?)(\d{2}?)(\d+).*"
                    }
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


def get_patients_from_folders(folder) -> List[Tuple[str, List[str]]]:
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
    patients_to_return = []
    for key in patient_patterns:
        regex_term = re.compile(patient_patterns[key])
        directories = set([x.string[:x.regs[-1][1]] for x in map(regex_term.fullmatch, files_folders) if x])
        matches = [[key] + list(x[0]) for x in map(regex_term.findall, directories) if x]
        patients_to_return.extend(list(zip(directories, matches)))
    return patients_to_return


def convert_patient_info_to_df(patients):
    """ get data as dictionary """
    df_dict = defaultdict(list)
    patients_to_remove = []
    for patient in patients:
        data = patient[1]
        key = data[0]
        data = data[1:]

        if key == 'full_info':
            if data[5] == '':
                patients_to_remove.append(patient)
                continue
            df_dict['GRK Nummer'].append(None)
            df_dict['Name'].append(f"{data[0]}, {data[1]}")
            df_dict['Geburtsdatum'].append(f"{data[4]}.{data[3]}.{data[2]}")
            df_dict['OP-Datum'].append(f"{data[8]}.{data[7]}.{data[6]}")
            df_dict['Patient-ID'].append(f"{data[5]}")
        elif key == 'name_id_opdate_time':
            if data[2] == '' or data[2].find('_') >= 0:
                patients_to_remove.append(patient)
                continue
            df_dict['Patient-ID'].append(f"{data[2]}")
            df_dict['GRK Nummer'].append(None)
            df_dict['Name'].append(f"{data[0]}, {data[1]}")
            df_dict['Geburtsdatum'].append(f"{0}.{0}.{0}")  # no birthdate available
            df_dict['OP-Datum'].append(f"{data[5]}.{data[4]}.{data[3]}")

    [patients.remove(x) for x in patients_to_remove]
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

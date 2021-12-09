from openpyxl import load_workbook
import pandas as pd
from pathlib import Path
from typing_extensions import TypedDict
from typing import List, Tuple, OrderedDict
from collections import defaultdict
from xml.etree import ElementTree
from xmltodict import parse
from datetime import datetime
import os
import re

divider = os.path.sep if os.path.sep != "\\" else os.path.sep * 2
patient_patterns = {'full_info': ".*" + divider +"(.*?)_(.*?)_(\d{4}?)(\d{2}?)(\d{2}?)_(.*?)_(\d{4}?)(\d{2}?)(\d{2}?)_\d{6}(\d{2}?)(\d{2}?).*",
                    'name_id_opdate_time': ".*" + divider +"(.*?)_(.*?)_(.*?)_(\d{4}?)(\d{2}?)(\d{2}?)_(\d{2}?)(\d{2}?)(\d+).*"
                    }
EXCEL_HEADERS = ["GRK Nummer", "Name", "Geburtsdatum", "OP-Datum", "Patient-ID"]
XML_FILE_NAMES = ['Patient.xml', 'TreatmentInfo.xml']


PatientInfo = TypedDict('PatientInfo', {'GRK Nummer': str,
                                        'folder_dir': str,
                                        'Name': str,
                                        'Patient-ID': str,
                                        'Geburtsdatum': str,
                                        'OP-Datum': str})


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


def get_patients_from_folders(folder) -> List[PatientInfo]:
    """
    Given a folder, will search for all subfolders with the Storz folder naming and return them.
    Parameters
    ----------
    folder : folder to recursively search through

    Returns
    -------
        A list of tuples each with the folder path of the patient and a tuple of the extracted patient info parsed by regex

    """
    files_folders = [str(x) for x in Path(folder).rglob("*.xml")]
    xml_files = [p for p in files_folders if os.path.basename(p) in XML_FILE_NAMES]

    patients_to_return = []  # type: List[PatientInfo]
    for xml in xml_files:
        try:
            # Try and get data from xml
            tree = ElementTree.parse(xml)
            xml_data = tree.getroot()
            xmlstr = ElementTree.tostring(xml_data, encoding='utf-8', method='xml')
            patient_data = parse(xmlstr)  # type: OrderedDict
            key = list(patient_data.keys())[0]
        except ElementTree.ParseError as e:
            # Try and get data from folder name
            key = 'folder_name'
            pass

        patient = PatientInfo()
        patient['GRK Nummer'] = None
        if key == 'Patient':
            data = patient_data[key]
            if not data['PatID']:
                continue
            fname = data["PatFirstName"] if data["PatFirstName"] is None else data["PatFirstName"].lower()
            lname = data["PatName"] if data["PatName"] is None else data["PatName"].lower()
            patient['Name'] = f'{lname}, {fname}'
            if data['PatBirth']:
                bday = data['PatBirth'].split('.')
                patient['Geburtsdatum'] = f'{int(bday[0]):02d}.{int(bday[1]):02d}.{int(bday[2]):04d}'
            else:
                patient['Geburtsdatum'] = None
            patient['OP-Datum'] = data['ORDate'].split(' ')[0]
            patient['Patient-ID'] = data['PatID'].split('-')[0]
            patient['folder_dir'] = os.path.dirname(xml)
        elif key == 'ExportedTreatment':
            data = patient_data[key]['Patient']
            if not data['IDNumber']:
                continue
            fname = data["GivenName"] if data["GivenName"] is None else data["GivenName"].lower()
            lname = data["FamilyName"] if data["FamilyName"] is None else data["FamilyName"].lower()
            patient['Name'] = f'{lname}, {fname}'
            bday = data['BirthdayAsDate']
            if bday:
                day = datetime.fromisoformat(bday)
                patient['Geburtsdatum'] = f'{day.day:02d}.{day.month:02d}.{day.year:04d}'
            else:
                patient['Geburtsdatum'] = f'{0:02d}.{0:02d}.{0:04d}'
            op_day = patient_data[key]['Series']['ProcedureDateTime']
            op_day = datetime.fromisoformat(op_day.split('.')[0])
            patient['OP-Datum'] = f'{op_day.day}.{op_day.month}.{op_day.year}'
            patient['Patient-ID'] = data['IDNumber'].split('-')[0]
            patient['folder_dir'] = os.path.dirname(xml)
        elif key == 'folder_name':
            tmp_list = []
            for pattern in patient_patterns:
                regex_term = re.compile(patient_patterns[pattern])
                directories = set([x.string[:x.regs[-1][1]] for x in map(regex_term.fullmatch, [xml]) if x])
                matches = [[pattern] + list(x[0]) for x in map(regex_term.findall, directories) if x]
                tmp_list.extend(list(zip(directories, matches)))
                if matches:
                    break
            pat_dicts = []
            for pat in tmp_list:
                data = pat[1]
                key = data[0]
                data = data[1:]
                new_dict = PatientInfo()
                new_dict['GRK Nummer'] = None
                new_dict['folder_dir'] = pat[0]
                if key == 'full_info':
                    if data[5] == '':
                        continue
                    new_dict['Name'] = f"{data[0].lower()}, {data[1].lower()}"
                    new_dict['Geburtsdatum'] = f"{int(data[4]):02d}.{int(data[3]):02d}.{int(data[2]):04d}"
                    new_dict['OP-Datum'] = f"{data[8]}.{data[7]}.{data[6]}"
                    new_dict['Patient-ID'] = f"{data[5].split('-')[0]}"
                elif key == 'name_id_opdate_time':
                    if data[2] == '' or data[2].find('_') >= 0:
                        continue
                    new_dict['Patient-ID'] = f"{data[2].split('-')[0]}"
                    new_dict['Name'] = f"{data[0].lower()}, {data[1].lower()}"
                    new_dict['Geburtsdatum'] = f"{0:02d}.{0:02d}.{0:04d}"  # no birthdate available
                    new_dict['OP-Datum'] = f"{data[5]}.{data[4]}.{data[3]}"
                else:
                    continue
                pat_dicts.append(new_dict)
            patients_to_return.extend(pat_dicts)
        else:
            continue
        if 'Patient-ID' in patient:
            patients_to_return.append(patient)
    return patients_to_return


def convert_patient_info_to_df(patients):
    """ get data as dictionary """
    df_dict = defaultdict(list)
    for patient in patients:
        for key in patient.keys():
            df_dict[key].append(patient[key])
    df_dict.pop('folder_dir')
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

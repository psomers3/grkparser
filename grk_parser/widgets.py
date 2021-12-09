from PyQt6.QtWidgets import *
from PyQt6.QtCore import *
from grk_parser.qcopy import CopyFiles
from grk_parser.util import get_df_from_excel, get_patients_from_folders, convert_patient_info_to_df, write_dataframe_to_excel
import pandas as pd
import os
from pathlib import Path
import math
from filetype import image_match, video_match


class FolderSelector(QWidget):
    """
    Widget with a LineEdit and a button to the right for selecting either a folder or a file
    """
    def __init__(self, folder: bool = True):
        """
        :param folder: whether or not to ask for a folder or a file
        """
        super(FolderSelector, self).__init__()
        self.setLayout(QHBoxLayout())
        self.file = QLineEdit()
        self.file.setMinimumWidth(500)
        if folder:
            self.file.setPlaceholderText('C:\\\\Path\\to\\folder\\')
            self.find_btn = QPushButton('Select Folder')
        else:
            self.file.setPlaceholderText('C:\\\\Path\\to\\excel_file.xlsx')
            self.find_btn = QPushButton('Select File')

        self.folder = folder
        self.find_btn.clicked.connect(self.select_folder)
        self.layout().addWidget(self.file)
        self.layout().addWidget(self.find_btn)

    def select_folder(self):
        """
        Called when the button is pushed. Fills the LineEdit with the folder/file selected.
        :return:
        """
        if self.folder:
            file = QFileDialog.getExistingDirectory(self, "Select Directory")
        else:
            file = QFileDialog.getOpenFileName(filter="Excel (*.xlsx)")[0]
        self.file.setText(file)

    def get_path(self):
        """
        :return: the text in the LineEdit
        """
        return self.file.text()


class MainWindow(QMainWindow):
    start_copying = pyqtSignal()

    def __init__(self):
        super(MainWindow, self).__init__()
        self.setWindowTitle("GRK Data Synchronizer")
        self.widget = QWidget()
        self.widget.setLayout(QVBoxLayout())
        self.setCentralWidget(self.widget)
        self.form_layout = QFormLayout()
        self.form_layout.setSpacing(0)
        label = QLabel('   The highest level folder that will be recursively searched for patient folders with videos or image data.')
        label.setToolTip('Der Ordner auf höchster Ebene, der rekursiv nach Patientenordnern mit Videos oder Bilddaten durchsucht werden soll.')
        self.form_layout.addRow('Source Folder:', label)
        self.form = QWidget()
        self.form.setLayout(self.form_layout)
        self.widget.layout().addWidget(self.form)
        self.source_folder = FolderSelector()
        self.destination_folder = FolderSelector()
        self.form_layout.addRow("", self.source_folder)
        label = QLabel('   The folder that contains the GRK anonymized patient storage.')
        label.setToolTip('Der Ordner, der die anonymisierte Patientenablage des GRK enthält.')
        self.form_layout.addRow('Destination Folder:', label)
        self.form_layout.addRow('', self.destination_folder)
        self.excel_file = FolderSelector(folder=False)
        excel_text = QLabel()
        excel_text.setTextFormat(Qt.TextFormat.RichText)
        excel_text.setText('&nbsp;&nbsp; The excel file with the record of GRK number with patient ID. This excel file needs to have a sheet named "Patienten" where<br>&nbsp;&nbsp;&nbsp;&nbsp;the data is stored. See the example excel sheet.<b> The file cannot be open when running this script</b>.')
        excel_text.setToolTip('Die Exceldatei mit der Aufzeichnung der GRK-Nummer mit Patienten-ID. Diese Excel-Datei muss ein Blatt mit dem Namen "Patienten" enthalten, in dem die Daten gespeichert werden. Siehe das Beispiel-Excel-Blatt. <b>Die Datei kann nicht geöffnet werden, wenn dieses Skript ausgeführt wird</b>.')
        self.form_layout.addRow('Excel file:', excel_text)
        self.form_layout.addRow('', self.excel_file)
        self.buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Cancel | QDialogButtonBox.StandardButton.Ok)
        self.buttons.accepted.connect(self.start_processing)
        self.buttons.rejected.connect(self.close)
        self.widget.layout().addWidget(self.buttons)

        self.progress = QProgressBar()
        self.progress.hide()
        self.progress.setRange(0, 1000)

        self.copier = CopyFiles()
        self.copier_thread = QThread()
        self.copier.moveToThread(self.copier_thread)
        self.start_copying.connect(self.copier.start_copying)
        self.copier.percent_copied.connect(self.progress.setValue)
        self.copier.finished.connect(self.on_finish)
        self.copier_thread.start()
        self.combined_df = None  # type: pd.DataFrame

    def on_finish(self):
        """
        Called when the copying is finished.
        """
        self.progress.hide()
        write_dataframe_to_excel(dataframe=self.combined_df, excel_file=self.excel_file.get_path())
        self.buttons.setDisabled(False)

    def start_processing(self):
        """
        This function get the patient data from the excel sheet and adds any new data found in the source folder. Once
        this is done, the data copier is started.
        """
        source_folder = self.source_folder.get_path()
        destination = self.destination_folder.get_path()
        if source_folder == "" or destination == "" or self.excel_file.get_path() == "":
            return
        self.buttons.setDisabled(True)
        if not os.path.exists(destination):
            os.makedirs(destination)
        patients = get_patients_from_folders(source_folder)
        existing_data = get_df_from_excel(self.excel_file.get_path())
        as_dict = convert_patient_info_to_df(patients)
        new_df = pd.DataFrame.from_dict(as_dict)

        combined = pd.concat([existing_data, new_df], ignore_index=True).drop_duplicates(
            subset=["Patient-ID", "OP-Datum"])
        combined = combined[combined.Name.notnull()].reset_index(drop=True)

        max_grk_id = float(combined['GRK Nummer'].max())
        if math.isnan(max_grk_id):
            max_grk_id = 0

        folders_to_copy = []
        duplicated_patients = combined[combined['Patient-ID'].duplicated(keep=False) == True]
        duplicated_name_and_birth = combined[combined.duplicated(subset=['Name', 'Geburtsdatum'], keep=False) == True]

        for index, row in combined.iterrows():
            if pd.isnull(combined.at[index, 'GRK Nummer']):
                pat_id = combined.at[index, 'Patient-ID']
                op_date = combined.at[index, 'OP-Datum']
                bdate = combined.at[index, 'Geburtsdatum']
                pname = combined.at[index, 'Name']

                if duplicated_patients.isin({'Patient-ID': [pat_id]})["Patient-ID"].any():
                    existing_ids = duplicated_patients[duplicated_patients['Patient-ID'] == pat_id]
                    existing_id = existing_ids[existing_ids["GRK Nummer"].isnull() == False]
                    if existing_id.empty:
                        max_grk_id += 1
                        combined.at[index, 'GRK Nummer'] = max_grk_id
                        duplicated_patients = combined[combined['Patient-ID'].duplicated(keep=False) == True]
                    else:
                        combined.at[index, 'GRK Nummer'] = existing_id['GRK Nummer'].iat[0]
                elif duplicated_name_and_birth.isin({'Name': [pname], 'Geburtsdatum': [bdate]})['Name'].any():
                    print(pname)
                    existing_ids = duplicated_name_and_birth[duplicated_name_and_birth['Name'] == pname]
                    existing_id = existing_ids[existing_ids["GRK Nummer"].isnull() == False]
                    if existing_id.empty:
                        max_grk_id += 1
                        combined.at[index, 'GRK Nummer'] = max_grk_id
                        duplicated_name_and_birth = combined[combined.duplicated(subset=['Name', 'Geburtsdatum'], keep=False) == True]
                    else:
                        combined.at[index, 'GRK Nummer'] = existing_id['GRK Nummer'].iat[0]
                    pass
                else:
                    max_grk_id += 1
                    combined.at[index, 'GRK Nummer'] = max_grk_id
                    duplicated_patients = combined[combined['Patient-ID'].duplicated(keep=False) == True]
                    duplicated_name_and_birth = combined[combined.duplicated(subset=['Name', 'Geburtsdatum'], keep=False) == True]

                current_grk_id = combined.at[index, 'GRK Nummer']

                for patient in patients:
                    if patient['Patient-ID'] == pat_id and patient['OP-Datum'] == op_date:
                        folders_to_copy.append((f"grk_{int(current_grk_id):04d}", patient['OP-Datum'], patient['folder_dir']))
        src = []
        dst = []
        for new_data_to_copy in folders_to_copy:
            copy_dest = os.path.join(destination, new_data_to_copy[0])
            if not os.path.exists(copy_dest):
                os.makedirs(copy_dest)
            p = Path(new_data_to_copy[-1]).glob("**/*")
            files = [x for x in p if x.is_file()]
            src_files = [str(x) for x in files if image_match(x) is not None or video_match(x) is not None]
            src.extend(src_files)
            dst.extend([os.path.join(destination, new_data_to_copy[0], new_data_to_copy[1], x[len(new_data_to_copy[-1]) + 1:]).replace("\\", "/") for x in src_files])
        self.copier.set_files_to_copy(src, dst)
        self.combined_df = combined
        self.progress.show()
        self.start_copying.emit()




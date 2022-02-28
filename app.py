import pandas as pd
from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import QFileDialog
import sys
import os
from pathlib import Path
from datetime import datetime

pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)


class FindMatch(QtWidgets.QMainWindow):

    def __init__(self):
        super(FindMatch, self).__init__()
        uic.loadUi("Template//MainWindow.ui", self)
        # print(os.path.dirname(sys.argv[0]))
        self.pushButton.clicked.connect(self.biolab_file_location)
        self.pushButton_2.clicked.connect(self.reject_folder_location)
        self.pushButton_3.clicked.connect(self.matching_name)
        self.file_location = ''
        self.folder_location = ''
        self.append_data = []

    def biolab_file_location(self):
        self.pushButton_3.setEnabled(True)
        self.file_location = QFileDialog.getOpenFileName(
            parent=self,
            caption='select Biolab File',
            directory=os.path.dirname(sys.argv[0]))[0]
        self.textEdit.setText(self.file_location)

    def reject_folder_location(self):
        self.pushButton_3.setEnabled(True)
        self.folder_location = QFileDialog.getExistingDirectory(
            parent=self,
            caption='Select reject folder location',
            directory=os.path.dirname(sys.argv[0]))
        # print(folder_location)
        self.textEdit_2.setText(self.folder_location)

    def matching_name(self):
        self.pushButton_3.setEnabled(False)
        for root, directories, files in os.walk(self.folder_location):
            for file in files:
                if file.endswith(".xlsx") or file.endswith('.xls') or file.endswith('.XLS'):
                    file_dir = root + '\\' + file
                    # print(file_dir)
                    df = pd.read_excel(file_dir, index_col=False, skiprows=2)
                    # print(df['Sample Number'])
                    self.append_data.append(df)
        append_data = pd.concat(self.append_data, ignore_index=True)
        # print(Path(self.file_location).parent.absolute())
        parent_dir = Path(self.file_location).parent.absolute()
        append_data.to_excel(str(parent_dir) + '\\combine.xlsx', index=False)

        df_1 = pd.read_excel(self.file_location, index_col=False)
        final_format = ['Date', 'Sponsor Name', 'PO#', 'IP Number', 'USB ID', 'Sample ID', 'VT ID', 'Origin Site',
                       'Diagnosis Category', 'Histopathological Diagnosis', 'Tissue Type', 'Necrosis Percentage',
                       '% Viable Tumor', 'Tissue Acceptable For Study', 'Comment']

        final_format_real = ['Date', 'Sponsor Name', 'Purchase Order', 'Request Number', 'USB ID', 'Sample Number',
                             'Ventana ID', 'Tissue Type', 'Diagnosis Category', 'Histopathological Diagnosis',
                             'Tissue Type',
                             'Necrosis Percentage', '% Viable Tumor', 'Tissue Acceptable for Study', 'Comment'
                             ]

        empty = ['Date', 'Sponsor Name', 'USB ID']
        from_biolab = ['Purchase Order', 'Request Number', 'Sample Number', 'Ventana ID']
        from_append = ['Tissue Type', 'Diagnosis Category', 'Histopathological Diagnosis', 'Tissue Type',
                       'Necrosis Percentage',
                       '% Viable Tumor', 'Tissue Acceptable for Study', 'Comment']

        result = []
        all_result = []
        error_list = set()

        for item, row in append_data.iterrows():
            bio_row = df_1[df_1['Ventana ID'] == row['Ventana ID']]
            if len(bio_row) == 0:
                error_list.add(row['Ventana ID'])
                continue
            else:
                for i in final_format_real:
                    if i in from_biolab:
                        result.append(bio_row[i].item())
                    elif i in from_append:
                        # print(row[i])
                        result.append(row[i])
                    else:
                        # print("empty")
                        result.append('')

            all_result.append(result)
            result = []

        # print(all_result)

        final_result = pd.DataFrame(all_result, columns=final_format)
        final_result_missing = pd.DataFrame(error_list, columns=["Miss from US Biolabs list"])

        final_result_dir = str(parent_dir) + "\\" + str(datetime.now().date()) + 'finalResult.xlsx'
        final_result_missing_dir = str(parent_dir) + "\\" + str(datetime.now().date()) + 'MissingResult.xlsx'
        final_result.to_excel(final_result_dir, index=False)
        final_result_missing.to_excel(final_result_missing_dir, index=False)
        os.startfile(str(parent_dir) + "\\" + str(datetime.now().date()) + 'finalResult.xlsx')
        os.startfile(str(parent_dir) + "\\" + str(datetime.now().date()) + 'MissingResult.xlsx')


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    qt_app = FindMatch()
    qt_app.show()
    sys._excepthook = sys.excepthook

    def exception_hook(exctype, value, traceback):
        print(exctype, value, traceback)
        sys.excepthook(exctype, value, traceback)
        sys.exit(1)

    sys.excepthook = exception_hook

    app.exec_()

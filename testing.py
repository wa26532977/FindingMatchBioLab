import pandas as pd
from PyQt5 import QtWidgets, uic, QtGui
from PyQt5.QtWidgets import QApplication, QWidget, QComboBox, QPushButton, QFileDialog, QVBoxLayout
import sys
import os
from datetime import datetime


pd.options.display.max_columns = 999
pd.options.display.max_rows = 999
pd.set_option("display.precision", 6)

folder_location = r"C:\Users\GG\Desktop\Testing\Reject_files"

# create_date = os.path.getctime(folder_location + '\\' + "rejcted-samples (6).xlsx")
# print(datetime.fromtimestamp(create_date).strftime('%Y-%m-%d %H:%M:%S'))


append_data = []
for root, directories, files in os.walk(folder_location):
    for file in files:
        if file.endswith(".xlsx") or file.endswith('.xls') or file.endswith('.XLS'):
            file_dir = root + '\\' + file
            print(file_dir)
            df = pd.read_excel(file_dir, index_col=False, skiprows=2)
            # print(df['Sample Number'])
            append_data.append(df)

append_data = pd.concat(append_data, ignore_index=True)
append_data.to_excel(r'C:\Users\GG\Desktop\Testing\append1.xlsx', index=False)
print("final result:")
# print(append_data)

biolab_list_dir = r'C:\Users\GG\Desktop\Testing\US Biolabs 21 Feb 2022.xlsx'
df_1 = pd.read_excel(biolab_list_dir, index_col=False)
# print("biolab_list_dir")
# print(df_1[df_1['Ventana ID'] == 'VT0000436241'])
# print("append_data")
# print(append_data[append_data['Ventana ID'] == 'VT0000436241'])

final_format = ['Date', 'Sponsor Name',	'PO#', 'IP Number', 'USB ID', 'Sample ID', 'VT ID',	'Origin Site',
                'Diagnosis Category', 'Histopathological Diagnosis', 'Tissue Type', 'Necrosis Percentage',
                '% Viable Tumor', 'Tissue Acceptable For Study', 'Comment']

final_format_real = ['Date', 'Sponsor Name', 'Purchase Order', 'Request Number', 'USB ID', 'Sample Number',
                     'Ventana ID', 'Tissue Type', 'Diagnosis Category', 'Histopathological Diagnosis', 'Tissue Type',
                     'Necrosis Percentage', '% Viable Tumor', 'Tissue Acceptable for Study', 'Comment'
                     ]

empty = ['Date', 'Sponsor Name', 'USB ID']
from_biolab = ['Purchase Order', 'Request Number', 'Sample Number', 'Ventana ID']
from_append = ['Tissue Type', 'Diagnosis Category', 'Histopathological Diagnosis', 'Tissue Type', 'Necrosis Percentage',
               '% Viable Tumor', 'Tissue Acceptable for Study', 'Comment']

cant_find = []
result = []
all_result = []
first_run = True

error_list = set()
for item, row in append_data.iterrows():
    bio_row = df_1[df_1['Ventana ID'] == row['Ventana ID']]
    # print(len(bio_row))
    # print(bio_row)
    # print(row)
    # # bio_row = df_1[df_1['Ventana ID'] == '001']
    # print(bio_row)
    # result.append(bio_row['Purchase Order'])
    # try:
    #     bio_row = df_1[df_1['Ventana ID'] == row['Ventana ID']]
    # except:
    #     cant_find.append(row['Ventana ID'])
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

final_result = pd.DataFrame(all_result, columns=final_format)
final_result_missing = pd.DataFrame(error_list, columns=["Miss from US Biolabs list"])
print(final_result_missing)

final_result_dir = r'C:\Users\GG\Desktop\Testing\finalResult' + str(datetime.now().date()) + '.xlsx'
final_result_missing_dir = r'C:\Users\GG\Desktop\Testing\MissingResult' + str(datetime.now().date()) + '.xlsx'
final_result.to_excel(final_result_dir, index=False)
final_result_missing.to_excel(final_result_missing_dir, index=False)

print(all_result)
print(error_list)






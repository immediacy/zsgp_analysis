import pandas as pd
import PyPDF2
import os
from time import ctime, time
from usef import path_converter
import json

xl_filename = path_converter(input('input a path to xls/xlsx file: '))
sheet = input('input sheet name: ')
folder_path = path_converter(input('input a path to .pdf files folder: '))

print(ctime())
start_time = time()

if xl_filename == '':
    xl_filename = r"F:\Том 11.3 КАЦ.xlsx"
if folder_path == '':
    folder_path = r"F:\258-8ТГС"

result_folder_path = '\\'.join(folder_path.split('\\')[:-1]) + \
                     f'\\renamed'
if not os.path.isdir(result_folder_path):
    os.mkdir(result_folder_path)
pdf_dir = os.listdir(folder_path)

for i in range(len(pdf_dir) - 1, -1, -1):
    if '.pdf' not in pdf_dir[i]:

        pdf_dir.pop(i)

xl = pd.read_excel(xl_filename, sheet_name=sheet)
anchor = '№\nпп'
first_row = 'not_defined'
first_column = 'not_defined'

for i in range(5):
    if len(xl[xl[xl.columns[i]] == anchor].index):
        first_row = xl[xl[xl.columns[i]] == anchor].index[0]
        first_column = i
        break
if first_row == 'not_defined' or first_column == 'not_defined':
    print('First row and column is not_defined')
    first_row = input('Input first row ')
    first_column = input('Input first column ')
xl.columns = xl.iloc[first_row, :]

xl = xl.iloc[first_row + 2:, first_column:]
main_column = 'Номер страницы в книге ценообразовывающих документов'
xl_len = xl.shape[0]
main_column_number = list(xl.columns).index(main_column)


j_path = xl_filename[:-5] + 'page_names_dict.json'

if os.path.exists(j_path):
    with open(j_path, 'r') as j_handle:
        page_names_dict = json.load(j_handle)

else:
    page_names_dict = {'last_page': 1}
page_counter = page_names_dict['last_page']
paragraph_name = ''
xl['new_name'] = pd.NA
xl['replica_flag'] = pd.NA
xl.reset_index(drop=True, inplace=True)

def_list = list()
for i in range(xl_len):
    cell_val = xl[main_column].iloc[i]
    if pd.notna(cell_val):
        if 'Том' in str(cell_val):
            paragraph_name = cell_val.split(',')[0]

for row_num in range(xl_len):
    cell_value = xl[main_column].iloc[row_num]
    if pd.notna(cell_value):
        if cell_value not in page_names_dict.keys():
            if str(cell_value) + '.pdf' in pdf_dir:
                original_pdf_path = folder_path + '\\' + str(cell_value) + '.pdf'
                with open(original_pdf_path, 'rb') as file:
                    reader = PyPDF2.PdfReader(file)
                    pdf_pages_number = len(reader.pages)

                if pdf_pages_number == 1:
                    result_filename = f'{paragraph_name}, cтр.{page_counter}'
                    page_counter += 1
                else:
                    page_number = f'{page_counter}-{page_counter + pdf_pages_number - 1}'
                    result_filename = f'{paragraph_name}, cтр.{page_number}'
                    page_counter += pdf_pages_number

                result_pdf_path = f'{result_folder_path}\\{result_filename}.pdf'

                os.popen(f'copy "{original_pdf_path}" "{result_pdf_path}"')
                page_names_dict[cell_value] = result_filename
                xl.loc[row_num, 'new_name'] = result_filename
            elif str(cell_value) not in def_list:
                def_list.append(str(cell_value))
                print(f'{cell_value}.pdf doesn\'t exist')
            else:
                pass
        else:
            xl.loc[row_num, 'new_name'] = page_names_dict[cell_value]
            xl.loc[row_num, 'replica_flag'] = 'Replica'

page_names_dict['last_page'] = page_counter
xl.to_excel(f'{folder_path}\\rename{ctime().replace(':', '_')}.xlsx')

with open(j_path, 'w') as jw_handle:
    json.dump(page_names_dict, jw_handle)

defect_list_path = f'{result_folder_path}\\pdf_not_found.json'

with open(defect_list_path, 'w') as def_handle:
    json.dump(def_list, def_handle)
print(ctime())
print(f'time seconds {time() - start_time}')

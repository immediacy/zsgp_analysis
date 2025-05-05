import traceback
import json
import pandas as pd
from usef import path_converter
import os
import aspose.pdf as ap
from time import ctime

xls_path = path_converter(input('Введите путь excel таблицы:'))
# xls_sheet_name = input('Введите номера лисов:')
xls_sheet_name = ['Оборудование ']  #, 'Оборудование'
anchor_row = ""
anchor_column = ""
# anchor_row = input('Input anchor for the first row or press Enter: ')
# anchor_column = input('Input anchor for the last column or press Enter')
pdf_path = path_converter(input('Введите путь папки с .pdf файлами:'))

j_path = xls_path[:-5] + 'stamper_lost_and_found.json'

if os.path.exists(j_path):
    with open(j_path, 'r') as j_handle:
        lost_and_found = json.load(j_handle)
else:
    lost_and_found = list()
res_path = pdf_path + f'\\stamped' \
           + ctime().replace(':', '_')
try:
    os.mkdir(res_path)
except FileExistsError:
    print('Folder exist already')
    pass

if len(anchor_column) == 0:
    anchor_column = 'Номер страницы в книге ценообразовывающих документов'
if len(anchor_row) == 0:
    anchor_row = ["№\nпп", '№ пп']

pdf_dir = os.listdir(pdf_path)
# remove a non pdf names
for i in range(len(pdf_dir) - 1, -1, -1):
    if '.pdf' not in pdf_dir[i]:

        pdf_dir.pop(i)
df_list = list()
for sheet_name in xls_sheet_name:
    df_list.append(pd.read_excel(xls_path, sheet_name=sheet_name))
first_row = 13
first_column = 0
def stamp_position(page_):

    page_height = page_.art_box.height
    page_width = page_.art_box.width
    stamp_width = len(matter_list[0]) * 4
    return page_width - stamp_width - 50, page_height * 0.3

number_of_columns = 33

for df in range(len(df_list)):
    for item in range(5):
        for cell in anchor_row:
            if (df_list[df].iloc[:, item] == cell).sum():
                first_row = df_list[df][df_list[df]\
                                            .iloc[:, item] == cell]\
                                            .index[0]
                first_column = item
                break
    df_list[df] = df_list[df].iloc[first_row:, first_column:]
    columns_names = df_list[df].iloc[0, :].dropna()
    number_of_columns = columns_names.shape[0]
    df_list[df] = df_list[df].iloc[2:, :number_of_columns + 1]
    df_list[df].rename(columns=columns_names, inplace=True)


df = pd.concat(df_list)
df_na = df[df[anchor_column].notna()]
invoice_col = 'Гиперссылка на веб-сайт производителя/поставщика '
for column_num, column_name in enumerate(columns_names):
    if pd.notna(column_name) and 'Гиперссылка' in column_name:
        invoice_col = column_name
        break
    elif column_num == number_of_columns - 1:
        invoice_col = input('Input the column name for invoice number: ')

delivery_msg = '* без учета заготовительно-складских и транспортных расходов'
matter_list = ['Наименование производителя/поставщика',
               'КПП организации ',
               'ИНН организации ',
               'Юридический адрес',
               'Телефон',
               'Электронная почта']
len_df = df_na.shape[0]
df_na.fillna('-', inplace=True)
delivery_disclaimer_flag = False

df_na[anchor_column] = df_na[anchor_column].astype(str)
for i in pdf_dir:

    sub_array = df_na[df_na[anchor_column] == i[0:-4]]

    if sub_array.shape[0] > 0:
        invoice_cell = sub_array.loc[sub_array.index[0], invoice_col]
        if i in lost_and_found:
            lost_and_found.remove(i)
        else:
            lost_and_found.append(i)
            continue
    if sub_array.shape[0] > 0 and 'http' in invoice_cell:
        # Open document

        try:
            document = ap.Document(pdf_path + '\\' + i)
        except RuntimeError:
            traceback.print_exc()
            print(pdf_path + '\\' + i)
            continue
        except Exception:
            traceback.print_exc()
            continue
        # Get particular page
        doc_len = len(document.pages)
        for j in range(1, doc_len + 1):
            page = document.pages[j]
            stamp_pos = stamp_position(page)

            req_list = list(sub_array.loc[sub_array.index[0], matter_list])
            to_str_list = list()
            buffer = req_list[1]
            req_list[1] = req_list[2]
            req_list[2] = buffer
            req_list[1] = 'ИНН ' + str(req_list[1])
            req_list[2] = 'КПП ' + str(req_list[2])
            req_list[3] = 'Юр. адрес: ' + str(req_list[3])

            for el in req_list:
                to_str_list.append(str(el))
            stamp = '\r\n'.join(to_str_list)
            date = sub_array.loc[sub_array.index[0], "Дата"]
            # Create text fragment
            text_fragment = ap.text.TextFragment(stamp)
            text_fragment.position = ap.text\
                .Position(stamp_pos[0], stamp_pos[1])

            # Set text properties
            text_fragment.text_state.font_size = 7
            text_fragment.text_state.font = ap.text\
                .FontRepository.find_font("Tahoma")
            text_fragment.text_state.background_color = ap.Color.transparent
            text_fragment.text_state.foreground_color = ap.Color.black
            page.paragraphs.add(text_fragment)

            # Create second text fragment
            text_fragment_1 = ap.text.TextFragment(date)
            text_fragment_1.position = ap.text\
                .Position(7, page.art_box.height - 30)

            # Set text properties
            text_fragment_1.text_state.font_size = 5
            text_fragment_1.text_state.font = ap.text\
                .FontRepository.find_font("Tahoma")
            text_fragment_1.text_state.background_color = ap.Color.white
            text_fragment_1.text_state.foreground_color = ap.Color.black
            if delivery_disclaimer_flag:

                # delivery disclaimer

                delivery_disclaimer = ap.text\
                    .TextFragment(delivery_msg)
                delivery_disclaimer.position = ap.text\
                    .Position(stamp_pos[0], stamp_pos[1] + 50)

                # Set text properties
                delivery_disclaimer.text_state.font_size = 5
                delivery_disclaimer.text_state.font = ap.text\
                    .FontRepository.find_font("Tahoma")
                delivery_disclaimer.text_state.background_color = ap.Color\
                    .light_gray
                delivery_disclaimer.text_state.foreground_color = ap.Color\
                    .black

            # signature secret
            signature = ap.text\
                .TextFragment('created by the monitoring department')
            signature.position = ap.text.Position(0, 0)

            # Set text properties
            signature.text_state.font_size = 5
            signature.text_state.font = ap.text\
                .FontRepository.find_font("Tahoma")
            signature.text_state.background_color = ap.Color.transparent
            signature.text_state.foreground_color = ap.Color.transparent
            # Create TextBuilder object
            builder = ap.text.TextBuilder(page)

            # Append the text fragment to the PDF page

            builder.append_text(text_fragment_1)
            if delivery_disclaimer_flag:
                builder.append_text(delivery_disclaimer)
            builder.append_text(signature)
    elif (sub_array.shape[0] > 0 and ('КП' in invoice_cell or
          'Счет' in invoice_cell or
          'Счёт' in sub_array.loc[sub_array.index[0], invoice_col])):
        # Open document
        document = ap.Document(pdf_path + '\\' + i)

        # Get particular page
        doc_len = len(document.pages)
        if doc_len > 4:
            doc_len = 4
        for j in range(1, doc_len + 1):
            page = document.pages[j]
            stamp_pos = stamp_position(page)
            req_list = list(sub_array.loc[sub_array.index[0], matter_list])
            to_str_list = list()
            for el in req_list:
                to_str_list.append(str(el))
            stamp = '\r\n'.join(to_str_list)
            date = sub_array.loc[sub_array.index[0], "Дата"]

            # signature secret
            signature = ap.text.TextFragment('created by the monitoring department')
            signature.position = ap.text.Position(0, 0)

            # Set text properties
            signature.text_state.font_size = 5
            signature.text_state.font = ap.text.FontRepository.find_font("Tahoma")
            signature.text_state.background_color = ap.Color.white
            signature.text_state.foreground_color = ap.Color.white
            # Create TextBuilder object
            builder = ap.text.TextBuilder(page)

            # Append the text fragment to the PDF page
            builder.append_text(signature)
    else:
        continue
    # Save resulting PDF document.
    document.save(res_path + '\\' + str(i))
with open(j_path, 'w') as jw_handle:
    json.dump(lost_and_found, jw_handle)

import datetime
import traceback
import json
import pandas as pd
from usef import path_converter
import os
import pymupdf as mu
from time import ctime, time
from PIL import Image, ImageFont, ImageDraw
import tkinter as tk
import numpy as np

# collect all input data through paths
page_number_change_flag = True
# xls_sheet_name = input('Введите номера лисов:')

xls_path = ''
# define and process a list of pdf files names
pdf_path = path_converter(input('Введите путь целевой папки:'))
pdf_dir = os.listdir(pdf_path)
for file in pdf_dir:
    if 'xls' in file:
        xls_path = f'{pdf_path}\\{file}'
        break
if xls_path == '':
    xls_path = path_converter(input('Введите путь excel таблицы:'))
# remove a non pdf names
for i in range(len(pdf_dir) - 1, -1, -1):
    if '.pdf' not in pdf_dir[i]:
        pdf_dir.pop(i)
xls_file = pd.ExcelFile(xls_path)
xls_sheet_list = xls_file.sheet_names

input_message = 'Для выбора листа '

for sh_num, sh_name in enumerate(xls_sheet_list):
    input_message = input_message + f'\n{sh_name}\tвведите {sh_num} '

xls_sheet_name = xls_sheet_list[int(input(input_message))]
anchor_row = ""
anchor_column = ""
# anchor_row = input('Input anchor for the first row or press Enter: ')
# anchor_column = input('Input anchor for the last column or press Enter')

start_time = time()
paragraph_delimiter = '_'
# define an anchor cell

if len(anchor_row) == 0:
    anchor_row = ["№\nпп", "№ \nпп", "№\n пп", "№ \n пп", '№ пп', '№ п.п.']

# load xls file
df = pd.read_excel(xls_path, sheet_name=xls_sheet_name)

result_folder_path = '\\'.join(pdf_path.split('\\')[:-1]) + \
                     f'\\renamed {pdf_path.split('\\')[-1]}'
if not os.path.isdir(result_folder_path):
    os.mkdir(result_folder_path)
# load the lost and found json file if it doesn't exist
j_path_lost = result_folder_path + '\\stamper_lost_and_found.json'

if os.path.exists(j_path_lost):
    with open(j_path_lost, 'r') as j_handle:
        lost_and_found = json.load(j_handle)
else:
    lost_and_found = list()

first_row = 13
first_column = 0
number_of_columns = 33

# define anchor cell position
for item in range(5):
    for cell in anchor_row:
        if (df.iloc[:, item] == cell).sum():
            first_row = df.iloc[:, item][df.iloc[:, item] == cell] \
                .index[0]
            first_column = item
            break
    if first_column != 0 or first_row != 13:
        break

df = df.iloc[first_row:, first_column:]
columns_names = df.iloc[0, :]
number_of_columns = columns_names.shape[0]
df = df.iloc[2:, :number_of_columns + 1]
df.rename(columns=columns_names, inplace=True)
# specify important column names
matter_dict = {
    'Наименование производителя': 0,
    'ИНН': 0,
    'КПП': 0,
    'адрес': 0,
    'Телефон': 0,
    'почта': 0,
    'Гиперссылка': 0,
    'mail': 0,
    'Номер страницы': 0
}
meta_list = list(matter_dict.keys())
font_file = r"C:\price_book\ofont.ru_Arial.ttf"  # r"C:\price_book\ofont.ru_Tahoma.ttf"
invoice_col = 'Гиперссылка на веб-сайт производителя/поставщика '
for column_num, column_name in enumerate(columns_names):
    if pd.notna(column_name):
        for beacon in meta_list:
            if beacon in column_name:
                matter_dict[beacon] = column_name
                break

if matter_dict['Гиперссылка'] == 0:
    invoice_col = input('Input the column name for invoice number: ')
else:
    invoice_col = matter_dict['Гиперссылка']
if matter_dict['Номер страницы'] == 0:
    anchor_column = 'Номер страницы в книге ценообразовывающих документов'
else:
    anchor_column = matter_dict['Номер страницы']
stamp_keys = [
    'Наименование производителя',
    'ИНН',
    'КПП',
    'адрес',
    'Телефон',
    'почта',
    'mail'
]
delivery_msg = '* без учета заготовительно-складских и транспортных расходов'
matter_list = list()
for key, val in matter_dict.items():
    if key in stamp_keys and val != 0:
        matter_list.append(val)

df_len = df.shape[0]
df[anchor_column] = df[anchor_column].astype(str)
j_path = result_folder_path + '\\page_names_dict.json'

if os.path.exists(j_path):
    with open(j_path, 'r') as j_handle:
        page_names_dict = json.load(j_handle)

else:
    page_names_dict = {'last_page': 1}
page_counter = page_names_dict['last_page']


# white space parser
def parser(width,
           height,
           start_,
           box: np.array):
    diagonal = int((width ** 2 + height ** 2) ** 0.5)
    width_counter = width / diagonal
    height_counter = height / diagonal
    # if standard position is right down corner
    # explore the stamp square from top left corner
    if start_[0] != 0 and start_[1] != 0:
        start = [start_[0], start_[1]]
        # check first condition
        for i_first in range(diagonal):
            x_first = start[1] + int(i_first * width_counter)
            y_first = start[0] + int(i_first * height_counter)

            if np.all(box[y_first][x_first] == 255):
                continue
            else:
                return y_first, x_first, False

        # check second condition
        for i_second in range(width * 2):
            x_second = start[1] + i_second % width
            y_second = start[0] + (i_second // width) * height - 1
            if np.all(box[y_second][x_second] == 255):
                continue
            else:
                return y_second, x_second, False

        # check third condition
        for i_third in range(height * 2):
            x_third = start[1] + (i_third // height) * width - 1
            y_third = start[0] + i_third % height
            if np.all(box[y_third][x_third] == 255):
                continue
            else:
                return y_third, x_third, False

        # check fourth condition
        for i_fourth in range(int(height * width / 5)):
            x_fourth = start[1] + i_fourth * 5 % width
            y_fourth = start[0] + (i_fourth * 5 // width)
            if np.all(box[y_fourth][x_fourth] == 255):
                continue
            else:
                return y_fourth, x_fourth, False
        return start[0], start[1], True
    # if standard position is right up corner
    # explore the stamp square from bottom left corner
    elif start_[0] == 0 and start_[1] != 0:
        start = [start_[0] + height, start_[1]]
        # check first condition
        for i_first in range(diagonal):
            x_first = start[1] + int(i_first * width_counter)
            y_first = start[0] - int(i_first * height_counter)

            if np.all(box[y_first][x_first] == 255):
                continue
            else:
                return y_first, x_first, False

        # check second condition
        for i_second in range(width * 2):
            x_second = start[1] + i_second % width
            y_second = start[0] - (i_second // width) * height - 1
            if np.all(box[y_second][x_second] == 255):
                continue
            else:
                return y_second, x_second, False

        # check third condition
        for i_third in range(height * 2):
            x_third = start[1] + (i_third // height) * width - 1
            y_third = start[0] - i_third % height
            if np.all(box[y_third][x_third] == 255):
                continue
            else:
                return y_third, x_third, False

        # check fourth condition
        for i_fourth in range(int(height * width / 5)):
            x_fourth = start[1] + i_fourth * 5 % width
            y_fourth = start[0] - (i_fourth * 5 // width)
            if np.all(box[y_fourth][x_fourth] == 255):
                continue
            else:
                return y_fourth, x_fourth, False
        return start[0], start[1], True
    elif start_[0] != 0 and start_[1] == 0:
        start = [start_[0], start_[1] + width]
        # check first condition
        for i_first in range(diagonal):
            x_first = start[1] - int(i_first * width_counter)
            y_first = start[0] + int(i_first * height_counter)

            if np.all(box[y_first][x_first] == 255):
                continue
            else:
                return y_first, x_first, False

        # check second condition
        for i_second in range(width * 2):
            x_second = start[1] - i_second % width
            y_second = start[0] + (i_second // width) * height - 1
            if np.all(box[y_second][x_second] == 255):
                continue
            else:
                return y_second, x_second, False

        # check third condition
        for i_third in range(height * 2):
            x_third = start[1] - (i_third // height) * width - 1
            y_third = start[0] + i_third % height
            if np.all(box[y_third][x_third] == 255):
                continue
            else:
                return y_third, x_third, False

        # check fourth condition
        for i_fourth in range(int(height * width / 5)):
            x_fourth = start[1] - i_fourth * 5 % width
            y_fourth = start[0] + (i_fourth * 5 // width)
            if np.all(box[y_fourth][x_fourth] == 255):
                continue
            else:
                return y_fourth, x_fourth, False
        return start[0], start[1], True
    elif start_[0] == 0 and start_[1] != 0:
        start = [start_[0] + height, start_[1] + width]
        # check first condition
        for i_first in range(diagonal):
            x_first = start[1] - int(i_first * width_counter)
            y_first = start[0] - int(i_first * height_counter)
            if np.all(box[y_first][x_first] == 255):
                continue
            else:
                return y_first, x_first, False

        # check second condition
        for i_second in range(width * 2):
            x_second = start[1] - i_second % width
            y_second = start[0] - (i_second // width) * height - 1
            if np.all(box[y_second][x_second] == 255):
                continue
            else:
                return y_second, x_second, False

        # check third condition
        for i_third in range(height * 2):
            x_third = start[1] - (i_third // height) * width - 1
            y_third = start[0] - i_third % height
            if np.all(box[y_third][x_third] == 255):
                continue
            else:
                return y_third, x_third, False

        # check fourth condition
        for i_fourth in range(int(height * width / 5)):
            x_fourth = start[1] - (i_fourth * 5 % width)
            y_fourth = start[0] - (i_fourth * 5 // width)
            if np.all(box[y_fourth][x_fourth] == 255):
                continue
            else:
                return y_fourth, x_fourth, False
        return start[0], start[1], True


# parser manager
def white_space(width,
                height,
                image_handle,
                ws_cell_value="not defined",
                pos='right_up',
                direction='left',
                part=0.3):
    start_point = [0, 0]
    img_size = image_handle.size
    img_list = list(image_handle.getdata())
    img_array = np.array(img_list)
    img_array = img_array.reshape(img_size[1], img_size[0], 3)
    if 'down' in pos:
        start_point[0] = img_size[1] - height
        down = False
    else:
        start_point[0] = 0
        down = True
    if 'right' in pos:
        start_point[1] = img_size[0] - width
        right = True
    else:
        start_point[1] = 0
        right = False
    reserved_start = [start_point[1], start_point[0]]
    looking_white_space = True
    if down and right:
        space_bounds = [start_point[0] - start_point[0] * part,
                        0]
    elif not down and right:
        space_bounds = [start_point[0] + start_point[0] * part,
                        0]
    else:
        space_bounds = [0, 0]
    while looking_white_space:
        parce_result = parser(width,
                              height,
                              start_point,
                              img_array)

        if direction == 'left':
            first_condition = parce_result[1] - width > space_bounds[1]
            second_condition = parce_result[2] - height > space_bounds[0]
            first_start = [start_point[0], parce_result[1] - width]
            second_start = [space_bounds[0] - height, parce_result[1] - width]
        else:
            first_condition = parce_result[2] - height > space_bounds[0]
            second_condition = parce_result[1] - width > space_bounds[1]
            first_start = [parce_result[2] - height, start_point[1]]
            second_start = [space_bounds[0] - height, parce_result[1] - width]

        if parce_result[0]:
            return [start_point[1], start_point[0] + 20]
        elif first_condition:
            start_point[0] = first_start[0]
            start_point[1] = first_start[1]
            continue
        elif second_condition:
            start_point[0] = second_start[0]
            start_point[1] = second_start[1]
            continue
        else:
            print(f"end of options{ws_cell_value}")
            return reserved_start


paragraph_name = 'Том ИОС1.9'
delivery_disclaimer_flag = False
delivery_msg = '* без учета заготовительно-складских и транспортных расходов'
def_list = list()
anchor_series = df[anchor_column]

# put the row delimiter
def wrap_text(text, width):
    """Переносит текст в зависимости от заданной ширины."""
    lines = []
    words = text.split()
    current_line = ""

    for word in words:
        # Проверка, поместится ли слово в текущую строку
        if len(current_line) + len(word) + 1 <= width:
            if current_line:  # Если текущая строка не пустая, добавляем пробел
                current_line += " "
            current_line += word
        else:
            # Добавляем текущую строку в список и начинаем новую строку
            lines.append(current_line)
            current_line = word

    # Добавляем последнюю строку, если она не пустая
    if current_line:
        lines.append(current_line)

    return "\n".join(lines)


# stamp preparing function
def prepare_stamp(req_list_df, stamp_size):
    req_list = list(req_list_df)
    to_str_list = list()
    stamp_length = round(stamp_size[0] / 6, 0)
    req_list[0] = wrap_text(req_list[0], stamp_length)
    req_list[1] = 'ИНН ' + str(req_list[1])
    req_list[2] = 'КПП ' + str(req_list[2])
    address = wrap_text('Юр. адрес: ' + str(req_list[3]),
                        stamp_length)

    # line delimiter
    req_list[3] = wrap_text(address, stamp_length)

    # make a string with the data for the pdf
    for el in req_list:
        to_str_list.append(str(el).strip())

    return '\n'.join(to_str_list)


for i in range(df_len):
    cell_val = anchor_series.iloc[i]
    if pd.notna(cell_val):
        if 'Том' in str(cell_val):
            paragraph_name = cell_val.split(paragraph_delimiter)[0]
            break

font = mu.Font(fontfile=font_file)
white_color = mu.pdfcolor['white']
for df_row in df.index:

    page_num_cell = df.loc[df_row, anchor_column]
    invoice_cell = df.loc[df_row, invoice_col]

    if (pd.notna(page_num_cell) and
            page_num_cell not in page_names_dict.keys() and
            page_num_cell not in lost_and_found):

        if (str(page_num_cell) + '.pdf' in pdf_dir or
                str(page_num_cell).replace(';', '_') + '.pdf' in pdf_dir):

            if str(page_num_cell).replace(';', '_') + '.pdf' in pdf_dir:
                original_pdf_path = pdf_path + '\\' + str(page_num_cell.replace(';', '_')) + '.pdf'
            else:
                original_pdf_path = pdf_path + '\\' + str(page_num_cell) + '.pdf'
            secret_text = 'created by the monitoring department'

            with open(original_pdf_path, 'rb') as file:
                reader = mu.open(original_pdf_path)
                pdf_pages_number = reader.page_count

            if 'http' in invoice_cell or 'ссылка' in invoice_cell:
                try:
                    for n_page in range(pdf_pages_number):
                        page_handler = reader.load_page(n_page)
                        page_size = np.ceil(page_handler.mediabox_size) \
                            .astype(np.int64).tolist()

                        pixels = page_handler.get_pixmap()

                        color_mode = "RGBA" if pixels.alpha else "RGB"

                        image_header = Image.frombytes(color_mode,
                                                       page_size,
                                                       pixels.samples)
                        stamp_size = [200, 120]
                        stamp_font = ('Arial', 12)
                        stamp = prepare_stamp(df.loc[df_row, matter_list],
                                              stamp_size)

                        stamp_position = white_space(stamp_size[0],
                                                     stamp_size[1],
                                                     image_header,
                                                     ws_cell_value=page_num_cell)

                        # get data has to be added to pdf
                        date = df.loc[df_row, "Дата"]

                        page_handler.insert_font(fontname=stamp_font[0],
                                                 fontbuffer=font.buffer)

                        stamp_handle = page_handler.insert_text(stamp_position,
                                                                stamp,
                                                                fontname=stamp_font[0],
                                                                fontfile=font_file,
                                                                fontsize=stamp_font[1])
                        if isinstance(date, datetime.datetime):
                            date = str(date)[:10]
                        date_handle = page_handler.insert_text([10, 30],
                                                               date,
                                                               fontname=stamp_font[0],
                                                               fontfile=font_file,
                                                               fontsize=stamp_font[1])

                        if delivery_disclaimer_flag:
                            delivery_handle = page_handler.insert_text([0, page_size[1]],
                                                                       delivery_msg,
                                                                       fontname=stamp_font[0],
                                                                       fontfile=font_file,
                                                                       fontsize=stamp_font[1])

                        secret_handle = page_handler.insert_text([0, page_size[1]],
                                                                 secret_text,
                                                                 fontname=stamp_font[0],
                                                                 fontfile=font_file,
                                                                 fontsize=5,
                                                                 color=white_color)


                except ValueError:
                    traceback.print_exc()
                    print(page_num_cell)
                    continue
                except Exception:
                    traceback.print_exc()
                    print(f'{page_num_cell} unexpected error')
                    continue
            elif ('КП' in invoice_cell or
                  'Счет' in invoice_cell or
                  'Счёт' in invoice_cell):

                for n_page in range(pdf_pages_number):
                    page_handler = reader.load_page(n_page)

                    page_size = np.ceil(page_handler.mediabox_size).tolist()
                    page_handler.insert_font(fontname='Tahoma',
                                             fontbuffer=font.buffer)
                    secret_handle = page_handler.insert_text([0, page_size[1]],
                                                             secret_text,
                                                             fontname='Tahoma',
                                                             fontfile=font_file,
                                                             fontsize=5,
                                                             color=white_color)

            else:
                print('http, КП, Счет not found in ', page_num_cell)
                continue
            # Save resulting PDF document.
            if pdf_pages_number == 1 and page_number_change_flag:
                result_filename = f'{paragraph_name}, cтр.{page_counter}'
                page_counter += 1
            elif not page_number_change_flag:
                result_filename = page_num_cell
            else:
                page_number = f'{page_counter}-{page_counter + pdf_pages_number - 1}'
                result_filename = f'{paragraph_name}, cтр.{page_number}'
                page_counter += pdf_pages_number

            page_names_dict[page_num_cell] = result_filename
            df.loc[df_row, 'new_name'] = result_filename
            result_pdf_path = f'{result_folder_path}\\{result_filename}.pdf'
            reader.ez_save(result_pdf_path)
        elif page_num_cell != 'nan':
            lost_and_found.append(str(page_num_cell) + '.pdf')
    elif pd.notna(page_num_cell) and page_num_cell in page_names_dict.keys():
        df.loc[df_row, 'new_name'] = page_names_dict[page_num_cell]
        df.loc[df_row, 'replica_flag'] = 'Replica'
    else:
        print(f'{page_num_cell} passed')

with open(j_path_lost, 'w') as jw_handle:
    json.dump(lost_and_found, jw_handle)

page_names_dict['last_page'] = page_counter
df.to_excel(f'{result_folder_path}\\rename{ctime().replace(':', '_')}.xlsx')

with open(j_path, 'w') as jw_handle:
    json.dump(page_names_dict, jw_handle)

print(ctime())
print(f'time seconds {time() - start_time}')

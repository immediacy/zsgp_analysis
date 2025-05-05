import pandas as pd
from usef import path_converter
import os
import time
import traceback

# a script for reformating a report from 1C

start_time = time.time()
x_path = r"C:\Users\User\Downloads\20250307 Анализ МПЗ Указчик 1 и 2.xlsx"
print(time.ctime())
g_path, g_name = path_converter(x_path, file_name=True)

# define a function to delete all useless fields from work dictionary
def clean_dict(dictionary):
    for key in dictionary.keys():
        if key != 'Заявка на МПЗ' and key != 'Заявка на ТМЦ':
            dictionary[key][1] = ''

# load the report
df = pd.read_excel(x_path)
# first row useful row number
iter_counter = 16
# redifine the iter_counter

if df.iloc[iter_counter, 0] != 'Заявка на МПЗ':

    iter_counter = df.iloc[:20, 0][df.iloc[:, 0].str.contains('Входящий номер')
    .fillna(False)].index[0] - 1
# make a new dataframe
data = df.iloc[iter_counter:, :].reset_index(drop=True)
# define all column names we're looking in the report
col_names = ['Наименование проекта',
             'Раздел проекта',
             'шифр проекта',
             'Заявка на МПЗ',
             'Заявка на ТМЦ',
             'Заказ поставщику',
             'Приходный ордер',
             'Наименование ТМЦ',
             'Наименование ТМЦ по счету',
             'Ед. изм.',
             'Ед. изм. замены',
             'Количество заявка на МПЗ',
             'Количество заказ поставщика',
             'Количество в приходном ордере',
             'Общее количество приходные ордера',
             'Общее количество заказы поставщикам',
             'Цена за единицу (фактическая), руб.',
             'НДС за ед., руб.',
             'Стоимость транспортных услуг Поставщика, руб.',
             'Общая стоимость с учетом НДС, руб.',
             'Бухгалтерские документы (с/ф, тн)',
             'Счет',
             'Поставщик контрагент',
             '1 платеж',
             '2 платеж',
             '3 платеж',
             '4 платеж',
             '5 платеж',
             'Дата поставки (планируемая ПТО)',
             'Базис поставки',
             'Текущее местонахождение тест',
             'Дата фактической поставки (по договору с Поставщиком)']

# to count amount of payment columns in the report and make a correct payment
# sum calculation define several vars
last_payment = '5 платеж'
payment_cols = []
columns_dict = {}
columns_list = []
payment_flag = True

# let's find all columns positions
# iterate all columns name list defined above
for column_name in col_names:
    # for each column name iterate header rows to correctly find the column position
    for n_row in range(10, 17):
        # make a short name for a series to explore
        temp_series = df.iloc[n_row, :]
        # payment series
        temp_ind = temp_series[temp_series.str.contains('платеж').fillna(False)]

        if temp_ind.shape[0] > 0:
            # take a last payment column position if temp_ind is not null
            last_payment = temp_ind[temp_ind.index[-1]]

        if temp_series[temp_series == column_name].shape[0] > 0:
            # the column name found write down the position
            temp_index = temp_series[temp_series == column_name].index[0]
            columns_dict[column_name] = [temp_index, '']

        if 'платеж' not in column_name:
            # all columns names goes to the columns_list except payment column
            columns_list.append(column_name)
        elif payment_flag:
            # payment column created as fixed string Total payment
            temp_index = temp_series[temp_series == column_name].index[0]
            payment_flag = False
            columns_list.append('Общий платеж')
            payment_cols.append(column_name)
            columns_dict['Общий платеж'] = [temp_index, '']
        else:
            # not reachable for any case
            payment_cols.append(column_name)
            print('Payment column name exception')

columns_list = list(set(columns_list))
payment_cols = list(set(payment_cols))
internal_number = ''
res_df = pd.DataFrame(columns=columns_list)
data_len = data.shape[0]
fill_flag = False
first_flag = True
# main loop iterate entire dataframe and reformat the cells
for row_num in range(data_len):
    try:
        # first cell in the row
        row_0 = data.loc[row_num, 0]
        # current row
        row_df = data.loc[row_num, :]
        if row_num + 1 < data_len:
            temp_zero = data.loc[row_num + 1, 0]
        else:
            temp_zero = ''
        if 'Заявка на МПЗ' in row_0:
            columns_dict['Заявка на МПЗ'] = row_0.split()[3] + " " + row_0.split()[6]
            continue
        elif 'Входящий номер' in row_0:
            temp_list = row_0.split()
            internal_number = temp_list[2] + " " + temp_list[4]
            columns_dict['Заявка на ТМЦ'][1] = internal_number
            continue
        elif 'Заказ поставщику' in row_0:
            columns_dict['Заказ поставщику'][1] = (row_0.split()[2] +
                                                   " " + row_0.split()[5])
            columns_dict['Общий платеж'][1] = row_df[
                        columns_dict['1 платеж'][0]: columns_dict[last_payment][0] + 1].sum()

            columns_dict['Количество заказ поставщика'][1] = row_df[
                columns_dict['Количество заказ поставщика'][0]]
            columns_dict['Дата поставки (планируемая ПТО)'][1] = row_df[
                columns_dict['Дата поставки (планируемая ПТО)'][0]]
            columns_dict['Цена за единицу (фактическая), руб.'][1] = row_df[
                columns_dict['Цена за единицу (фактическая), руб.'][0]]
            columns_dict['НДС за ед., руб.'][1] = row_df[columns_dict['НДС за ед., руб.'][0]]
            columns_dict['Стоимость транспортных услуг Поставщика, руб.'][1] = row_df[
                columns_dict['Стоимость транспортных услуг Поставщика, руб.'][0]]
            columns_dict['Общая стоимость с учетом НДС, руб.'][1] = row_df[
                columns_dict['Общая стоимость с учетом НДС, руб.'][0]]
            columns_dict['Поставщик контрагент'][1] = row_df[columns_dict['Поставщик контрагент'][0]]
            columns_dict['Счет'][1] = row_df[columns_dict['Счет'][0]]
            columns_dict['Текущее местонахождение тест'][1] = row_df[
                columns_dict['Текущее местонахождение тест'][0]]
            columns_dict['Бухгалтерские документы (с/ф, тн)'][1] = row_df[
                columns_dict['Бухгалтерские документы (с/ф, тн)'][0]]
            columns_dict['Дата фактической поставки (по договору с Поставщиком)'][1] = row_df[
                columns_dict['Дата фактической поставки (по договору с Поставщиком)'][0]]

            columns_dict['Ед. изм. замены'][1] = row_df[columns_dict['Ед. изм. замены'][0]]
            columns_dict['Наименование ТМЦ по счету'][1] = row_df[columns_dict['Наименование ТМЦ по счету'][0]]
            columns_dict['Базис поставки'][1] = row_df[columns_dict['Базис поставки'][0]]
            if (internal_number in temp_zero or
                    'Заказ поставщику' in temp_zero or
                    'Заявка на МПЗ' in temp_zero):
                fill_flag = True

        elif 'Приходный ордер' in row_0:
            temp_order_num = row_df[columns_dict['Приходный ордер'][0]].split()
            columns_dict['Приходный ордер'][1] = temp_order_num[2] + " " + temp_order_num[4]
            columns_dict['Количество в приходном ордере'][1] = row_df[columns_dict['Количество в приходном ордере'][0]]

            if (internal_number in temp_zero or
                    'Заказ поставщику' in temp_zero or
                    'Заявка на МПЗ' in temp_zero or
                    'Приходный ордер' in temp_zero):
                fill_flag = True

        elif internal_number in row_0:
            columns_dict['Количество заявка на МПЗ'][1] = row_df[columns_dict['Количество заявка на МПЗ'][0]]
            columns_dict['Общее количество приходные ордера'][1] = row_df[columns_dict['Общее количество приходные ордера'][0]]
            columns_dict['Наименование проекта'][1] = row_df[columns_dict['Наименование проекта'][0]]
            columns_dict['Раздел проекта'][1] = row_df[columns_dict['Раздел проекта'][0]]
            columns_dict['шифр проекта'][1] = row_df[columns_dict['шифр проекта'][0]]
            columns_dict['Наименование ТМЦ'][1] = row_df[columns_dict['Наименование ТМЦ'][0]]
            columns_dict['Ед. изм.'][1] = row_df[columns_dict['Ед. изм.'][0]]
            columns_dict['Общее количество заказы поставщикам'][1] = row_df[
                columns_dict['Общее количество заказы поставщикам'][0]]
            if not pd.isna(temp_zero) and internal_number in temp_zero:
                fill_flag = True
        else:
            continue
        # fill the data
        if fill_flag:
            temp_dict = columns_dict.copy()
            for col in payment_cols:
                temp_dict.pop(col, None)
            temp_df = pd.DataFrame(data=temp_dict)
            if first_flag:
                res_df = temp_df.iloc[1:2, :]
                first_flag = False
            else:
                res_df = pd.concat([res_df, temp_df.iloc[1:2, :]]).reset_index(drop=True)
            fill_flag = False
            if (internal_number in temp_zero or
                    'Заявка на МПЗ' in temp_zero):
                clean_dict(columns_dict)

    except IndexError:
        traceback.print_exc()
        print(row_num)

    except:
        traceback.print_exc()
        print(row_num)

c_time = time.strftime('%H_%M_%S')
new_filename = f' {c_time}.'.join(g_name.split('.'))
# upload the results to Excel
res_df.to_excel(os.path.join(g_path, new_filename))
print(f'{time.time() - start_time} seconds to run')

import pandas as pd
import openpyxl as ox
import numpy as np

ex_path = r"C:\Users\User\Downloads\Таблица по регламенту Камчатка.xlsx"
report_path = r"C:\Users\User\Downloads\Отчет 720.xlsx"

ndf = pd.read_excel(ex_path)
res_df = ndf.copy()
ndf.drop(range(6), axis=0, inplace=True)
ndf.columns = ndf.iloc[0]
ndf.drop([6, 7], axis=0, inplace=True)
ndf = ndf[ndf['ЗаявкаНаименование'].notna()]
rep = pd.read_excel(report_path)
first_rep_row = rep[rep['Unnamed: 0'] == 'Заявка на МПЗ'].index[0]
rep.drop(range(10), axis=0, inplace=True)

# define a dict to write down all changes
update_dict = {}

order_header_dict = ['Заявка на МПЗ.Комментарий',
                     'Заявка на МПЗ вн. номер и вн. дата',
                     'Заявка на МПЗ']

first_rep_row_values = first_rep_row + 5

# select the rows that contain the 'Заявка на МПЗ' string
selected_dict = rep[rep['Unnamed: 0'].str.contains('Заявка на МПЗ')]['Unnamed: 0'].head(5).to_dict()

# first_rep_row_values defining loop
if len(selected_dict) > 3:
    for r_num, r_value in selected_dict.items():
        if r_value in order_header_dict:
            continue
        else:
            first_rep_row_values = r_num
            break

rdf = rep.drop(range(first_rep_row, first_rep_row_values))
# ndf.loc[:, 'Статус оплаты'] = ndf['Статус оплаты'].fillna('nan')
item_name_col_name = 'Наименование материально-технических ресурсов'
for row in ndf.index:
    # get the current row from the reglament table
    temp_ndf_row = ndf.loc[row, :]
    # get the income order number
    temp_income_num = temp_ndf_row['Номер и дата заявки УСМР']
    # and get the item name
    temp_item_name = temp_ndf_row[item_name_col_name]
    # get temporary window indexes where the income num from reglament table
    # is identical to income num in the report
    r_window = rep[rep['Unnamed: 0'] == temp_income_num]
    if r_window.shape[0]:
        r_window_index = r_window.index
    else:
        continue
    # temp df which contains the income numbers we're looking
    income_window = rep.loc[r_window_index[0]: r_window_index[-1], :]

    item_row = income_window[income_window['Unnamed: 8'] == temp_item_name]
    if item_row.shape[0] == 0:
        update_dict[(row, item_name_col_name)] = 'not found'
        continue

    item_start_index = item_row.index[0]

    # temp report df where
    temp_item_df = income_window.loc[item_start_index:]
    item_end_index = item_start_index
    temp_item_df = temp_item_df.drop(temp_item_df.index[0], axis=0)

    for temp_row in temp_item_df.index:
        if (temp_item_df.loc[temp_row, 'Unnamed: 0'] == temp_income_num or
                order_header_dict[2] in temp_income_num):
            break
        else:
            item_end_index = temp_item_df.loc[temp_row].name
    # a partial df with name from report df
    item_rep_df = income_window.loc[item_start_index: item_end_index]
    comparison_dict_request = {'Раздел проекта': 'Unnamed: 6',
                               'Шифр проекта': 'Unnamed: 4',
                               'Ответственный от УМТО': 'Unnamed: 17',
                               'Наименование материально-технических ресурсов': 'Unnamed: 8',
                               'Ед.изм.': 'Unnamed: 9',
                               'Кол-во': 'Unnamed: 11',
                               'Дата поставки на объект указанная в заявке УСМР': 'Unnamed: 27'}
    for key, val in comparison_dict_request.items():
        temp_rep_cell = item_rep_df.loc[item_start_index, val]
        temp_ndf_cell = temp_ndf_row.loc[key]
        if not pd.isnull(temp_rep_cell) and temp_ndf_cell != temp_rep_cell:
            res_df.loc[row, key] = temp_rep_cell
            update_dict[(row, key)] = temp_rep_cell

    if item_rep_df.shape[0] == 1:
        continue
    # supplier order data insertion
    so_df = item_rep_df[item_rep_df['Unnamed: 0'].str.contains('Заказ поставщику')]

    total_price = so_df['Unnamed: 22'].sum()
    total_volume = so_df['Unnamed: 13'].sum()
    average_price = round(total_price / total_volume, 2)
    supplier_name = '\n'.join(so_df['Unnamed: 24'].unique())
    supply_basis = so_df['Unnamed: 24'].unique()
    clear_supply = supply_basis.copy()
    for num_1, basis_1 in enumerate(supply_basis):
        for num_2, basis_2 in enumerate(supply_basis[num_1 + 1:]):
            if (not isinstance(basis_1, str) or
                    not isinstance(basis_2, str)):
                continue
            temp_b_1 = basis_1.lower().replace('.', '').replace('г', '')
            temp_b_2 = basis_1.lower().replace('.', '').replace('г', '')
            if temp_b_1 in temp_b_2 or temp_b_2 in temp_b_1:
                clear_supply.remove(basis_1)
                break

    basis = '\n'.join(clear_supply)

    comparison_dict_order = {'Стоимость за ед. (руб.)': average_price,
                             'Общая стоимость': total_price,
                             'Наименование контрагента': supplier_name,
                             'Базис поставки': clear_supply}

    for key, val in comparison_dict_order.items():
        temp_rep_cell = val
        temp_ndf_cell = temp_ndf_row.loc[key]
        if not pd.isnull(temp_rep_cell) and temp_ndf_cell == temp_rep_cell:
            res_df.loc[row, key] = temp_rep_cell
            update_dict[(row, key)] = temp_rep_cell
    # supply data insertion
    sup_df = item_rep_df[item_rep_df['Unnamed: 0'].str.contains('Приходный ордер')]
    sup_value = sup_df['Unnamed: 16'].sum()

    res_df.loc[row, 'Поставлено'] = f'{round(sup_value / total_volume * 100, 2)}%'

import pandas as pd
import regex as re
from fontTools.ttLib.tables.O_S_2f_2 import table_O_S_2f_2

from sql_loads import drop_nan_columns

# define the paths
sale_doc = r"G:\Мой диск\Tasks\EUTP-105128 для МТ-7255984 Не заполняется подразделение в Е-смарт\РТЗ выгрузка за 2025 Реализации + талоны.xlsx"
invoice_doc = r"G:\Мой диск\Tasks\EUTP-105128 для МТ-7255984 Не заполняется подразделение в Е-смарт\РТЗ выгрузка за 2025 Счета + талоны.xlsx"
def rtz_counter(sale_doc_path, invoice_doc_path):
    # load the data
    sales = pd.read_excel(sale_doc_path)
    invoices = pd.read_excel(invoice_doc_path)
    # clean headers and rename columns
    sales = drop_nan_columns(sales)
    invoices = drop_nan_columns(invoices)
    # Merge the DFs
    merged = pd.merge(invoices, sales, how="outer", on='Ссылка')
    # drop NA
    merged_dn = merged.dropna(axis=0)

    # number of empty departments when ticket department is filled
    empty_dpts_1 = merged_dn[(~merged_dn['Подразделение_y'].str.contains('<Пустая ссылка: ')) &
                   (merged_dn['ПодразделениеОрганизации'].str.contains('<Пустая ссылка: ')) &
                   (merged_dn['ПодразделениеОрганизации1'].str.contains('<Пустая ссылка: '))]
    empty_dpts_1_num = empty_dpts_1.shape[0]
    # number of filled departments when ticket department is empty
    empty_dpts_2 = merged_dn[(merged_dn['Подразделение_y'].str.contains('<Пустая ссылка: ')) &
                  (~merged_dn['ПодразделениеОрганизации'].str.contains('<Пустая ссылка: ')) &
                  (~merged_dn['ПодразделениеОрганизации1'].str.contains('<Пустая ссылка: '))]
    empty_dpts_2_num = empty_dpts_2.shape[0]
    # number of empty departments when ticket department is empty
    empty_dpts_3 = merged_dn[(merged_dn['Подразделение_y'].str.contains('<Пустая ссылка: ')) &
              (merged_dn['ПодразделениеОрганизации'].str.contains('<Пустая ссылка: ')) &
              (merged_dn['ПодразделениеОрганизации1'].str.contains('<Пустая ссылка: '))]
    empty_dpts_3_num = empty_dpts_3.shape[0]
    # number of filled departments when ticket department is filled
    empty_dpts_4 = merged_dn[(~merged_dn['Подразделение_y'].str.contains('<Пустая ссылка: ')) &
              (~merged_dn['ПодразделениеОрганизации'].str.contains('<Пустая ссылка: ')) &
              (~merged_dn['ПодразделениеОрганизации1'].str.contains('<Пустая ссылка: '))]
    empty_dpts_4_num = empty_dpts_4.shape[0]
    empty_dpts_1.loc[:, 'date'] = pd.to_datetime(empty_dpts_1['ДатаСделкиРИЭС_x'], format='%d.%m.%Y')
    dates_1 = empty_dpts_1.groupby(empty_dpts_1.date.dt.to_period('M')).agg({'Подразделение_x': 'count'})
    empty_dpts_2.loc[:, 'date'] = pd.to_datetime(empty_dpts_2['ДатаСделкиРИЭС_x'], format='%d.%m.%Y')
    dates_2 = empty_dpts_2.groupby(empty_dpts_2.date.dt.to_period('M')).agg({'Подразделение_x': 'count'})
    empty_dpts_3.loc[:, 'date'] = pd.to_datetime(empty_dpts_3['ДатаСделкиРИЭС_x'], format='%d.%m.%Y')
    dates_3 = empty_dpts_3.groupby(empty_dpts_3.date.dt.to_period('M')).agg({'Подразделение_x': 'count'})
    empty_dpts_4.loc[:, 'date'] = pd.to_datetime(empty_dpts_4['ДатаСделкиРИЭС_x'], format='%d.%m.%Y')
    dates_4 = empty_dpts_4.groupby(empty_dpts_4.date.dt.to_period('M')).agg({'Подразделение_x': 'count'})
    print(f'total length {merged_dn.shape[0]}')
    print(f'number of empty departments when ticket department is filled {empty_dpts_1_num}\n'
          f'number of filled departments when ticket department is empty {empty_dpts_2_num}\n'
          f'number of empty departments when ticket department is empty {empty_dpts_3_num}\n'
          f'number of filled departments when ticket department is filled {empty_dpts_4_num}')
    return sales, invoices, dates_1, dates_2, dates_3, dates_4

# rtz = rtz_counter(sale_doc, invoice_doc)
# paying card operation
add_serv_pd_path = r"G:\Мой диск\Tasks\EUTP-105128 для МТ-7255984 Не заполняется подразделение в Е-смарт\Доп.услуги выгрузка 2025 Операции по платежной карте + ЭС.Доп.услуга.xlsx"
# cash order
add_serv_ip_path = r"G:\Мой диск\Tasks\EUTP-105128 для МТ-7255984 Не заполняется подразделение в Е-смарт\Доп.услуги выгрузка 2025 ПКО + ЭС.Доп.услуга.xlsx"
def add_serv(add_serv_pay_dev_path, add_serv_income_payment_path):
    # card operations
    pay_dev = pd.read_excel(add_serv_pay_dev_path)
    # card payment
    income_pay = pd.read_excel(add_serv_income_payment_path)
    pay_dev = drop_nan_columns(pay_dev)
    income_pay = drop_nan_columns(income_pay)
    empty_dpt_pd = pay_dev[(pay_dev.ПодразделениеОрганизации.str.contains('<Пустая ссылка:'))]
    empty_dpt_ip = income_pay[(income_pay.ПодразделениеОрганизации.str.contains('<Пустая ссылка:'))]
    print(f'total pd length {pay_dev.shape[0]}\n'
          f'total ip length {income_pay.shape[0]}\n'
          f'empty pd length {empty_dpt_pd.shape[0]}\n'
          f'empty ip length {empty_dpt_ip.shape[0]}')
    return empty_dpt_ip, empty_dpt_pd
# add_srv = add_serv(add_serv_pd_path, add_serv_ip_path)

pkpk_pko_path = (r"G:\Мой диск\Tasks"
                 r"\EUTP-105128 для МТ-7255984 Не заполняется подразделение в Е-смарт\ПКПК выгрузка за 2025 ПКО + ЭС.ПКПК.xlsx")
pkpk_payment_ops_path = (r"G:\Мой диск\Tasks"
                         r"\EUTP-105128 для МТ-7255984 Не заполняется подразделение в Е-смарт\ПКПК выгрузка за 2025 Операции по платежной карте + ЭС.ПКПК.xlsx")
pkpk_path = (r"G:\Мой диск\Tasks"
             r"\EUTP-105128 для МТ-7255984 Не заполняется подразделение в Е-смарт"
             r"\ПКПК выгрузка за 2025 ЭС.ПКПК.xlsx")
def pkpk_pko(pkpk_pko_path_, pkpk_po_path, pkpk_p):

    pkpk_pko = pd.read_excel(pkpk_pko_path_)
    pkpk_po = pd.read_excel(pkpk_po_path)
    pkpk = pd.read_excel(pkpk_p)
    pkpk_pko = drop_nan_columns(pkpk_pko)
    pkpk_po = drop_nan_columns(pkpk_po)
    pkpk = drop_nan_columns(pkpk)
    pkpk['doc_type'] = pkpk.ДокументСсылка.str.split('БЧ')
    pkpk['dt'] = pkpk.doc_type.map(lambda x: x[0])
    pkpk_pko_empty = pkpk_pko[pkpk_pko.ПодразделениеОрганизации.str.contains('<Пустая ссылка: ')]
    pkpk_po_empty = pkpk_po[pkpk_po.ПодразделениеОрганизации.str.contains('<Пустая ссылка: ')]
    print(f'empty departments counter for pko {pkpk_pko_empty.shape[0]}\n'
          f'empty departments counter for po {pkpk_po_empty.shape[0]}')
    return pkpk_pko, pkpk_po

# result_pkpk = pkpk_pko(pkpk_pko_path, pkpk_payment_ops_path, pkpk_path)

"""
cos_path = r"G:\Мой диск\Tasks\EUTP-105128 для МТ-7255984 Не заполняется подразделение в Е-смарт\ЦОС выгрузка за 2025 ЭС.ЦОС.xlsx"
cos = pd.read_excel(cos_path)
cos = drop_nan_columns(cos)
"""

# pko document oriented
pko_path_tmp = (r"G:\Мой диск\Tasks"
                r"\EUTP-105128 для МТ-7255984 Не заполняется подразделение в Е-смарт"
                r"\ПКО с назначением платежа.xlsx")
tsos_path_tmp = (r"G:\Мой диск\Tasks"
                 r"\EUTP-105128 для МТ-7255984 Не заполняется подразделение в Е-смарт"
                 r"\ЦОС выгрузка за 2025 ПКО_номерЗаявки(нуль+контрагент).xlsx")
pko_path_new = (r"G:\Мой диск\Tasks"
                r"\EUTP-105128 для МТ-7255984 Не заполняется подразделение в Е-смарт"
                r"\ПКО с заполненными подразделениями за 2025.xlsx")

def pko_tsos(pko_path, tsos_path):
    df = pd.read_excel(pko_path)
    df = drop_nan_columns(df)
    reg_exp = r"(\b\d{7,9}\b)"
    df['ticket'] = df.Коммент.str.extract(reg_exp)
    df.ticket = df.ticket.fillna('not_found')
    df_with_ticket_num = df[df.ticket != 'not_found']
    df_with_ticket_num.loc[:, 'ticket'] = df_with_ticket_num.ticket.astype('int32')
    tsos_df = pd.read_excel(tsos_path)
    tsos_df = drop_nan_columns(tsos_df)

    df3 = pd.merge(df_with_ticket_num, tsos_df, left_on='ticket', right_on='НомерЗаявкиРИЭС')


    return df3
df_2 = pko_tsos(pko_path_new, tsos_path_tmp)
df1 = pko_tsos(pko_path_tmp, tsos_path_tmp)
df1.to_excel("G:\Мой диск\Tasks\EUTP-105128 для МТ-7255984 Не заполняется подразделение в Е-смарт\ПКО с.xlsx")
print(f'the end shape{df1.shape}')
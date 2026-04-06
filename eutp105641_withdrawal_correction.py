# скрипт нужен чтобы выяснить сколько случаев нахождения дублей ИНН среди
# риелторов, оплативших обязательные услуги
import pandas as pd
from sql_loads import drop_nan_columns



contractors_path = (r"G:\Мой диск\Tasks"
                    r"\EUTP-105641 Корректировка документа списания с расчетного счета"
                    r"\Выгрузка Контрагенты с ИНН ФизЛицо.xlsx")
contractors_df = pd.read_excel(contractors_path)
contractors_df = drop_nan_columns(contractors_df)

contractors_etagi_path = (r"G:\Мой диск\Tasks"
                          r"\EUTP-105641 Корректировка документа списания с расчетного счета"
                          r"\Выгрузка Контрагенты все с ИНН Этажи.xlsx")
contractors_etagi_df = pd.read_excel(contractors_etagi_path)
contractors_etagi_df = drop_nan_columns(contractors_etagi_df)

withdraw_path = (r"G:\Мой диск\Tasks"
                 r"\EUTP-105641 Корректировка документа списания с расчетного счета"
                 r"\Документы списания + Назначение платежа.xlsx")
e_contracts_path = (r"G:\Мой диск\Tasks"
                    r"\EUTP-105641 Корректировка документа списания с расчетного счета"
                    r"\Выгрузка Договоры Е-Смарт с риелторами Новостройки+Вторичка.xlsx")

acc_contracts_path = (r"G:\Мой диск\Tasks"
                      r"\EUTP-105641 Корректировка документа списания с расчетного счета"
                      r"\Выгрузка Договоры Этажи с риелторами номинальный счет.xlsx")

e_contracts_df = pd.read_excel(e_contracts_path)
e_contracts_df = drop_nan_columns(e_contracts_df)
e_contracts_df_clean = e_contracts_df[e_contracts_df.ВладелецИНН != '<Пустая строка>']

acc_contracts_df = pd.read_excel(acc_contracts_path)
acc_contracts_df = drop_nan_columns(acc_contracts_df)
acc_contracts_df_clean = acc_contracts_df[acc_contracts_df.ВладелецИНН != '<Пустая строка>']

withdraw_docs = pd.read_excel(withdraw_path)
withdraw_docs = drop_nan_columns(withdraw_docs)

withdraw_docs['temp'] = withdraw_docs.НазначениеПлатежа.str.split('ИНН').map(lambda x: x[1])
withdraw_docs['temp'] = withdraw_docs.temp.str.split('Сумма').map(lambda x: x[0])
withdraw_docs['contractor'] = withdraw_docs.temp.str.strip()

withdraw_docs['temp'] = withdraw_docs.НазначениеПлатежа.str.split('ИНН')
withdraw_docs.temp = withdraw_docs.temp.map(lambda x: x[0])
withdraw_docs['temp'] = withdraw_docs.НазначениеПлатежа.str.split(' за ')
withdraw_docs.temp = withdraw_docs.temp.map(lambda x: x[1])
withdraw_docs.temp = withdraw_docs.temp.str.replace('ИП', "").str.strip()
withdraw_docs.temp = withdraw_docs.temp.str.split(' ')

withdraw_docs['last_name'] = withdraw_docs.temp.map(lambda x: x[0])
withdraw_docs['first_middle_name'] = withdraw_docs.temp.map(lambda x: " ".join(x[1:3]))
# withdraw_docs.drop('temp', axis=1, inplace=True)

with_contractors = withdraw_docs.contractor.unique()
active_contractors = e_contracts_df_clean[e_contracts_df_clean.ВладелецИНН.isin(with_contractors)]

more_than_one_contract = (active_contractors
                          .groupby(['ВладелецИНН', 'Наименование'], as_index=False)
                          .agg({'Владелец': 'count'})
                          .sort_values('Владелец'))
more_than_one_contract = more_than_one_contract[more_than_one_contract.Владелец > 1]
unique_more_than_one_contract = more_than_one_contract.drop_duplicates(subset=['ВладелецИНН'])

doubled_contract = withdraw_docs.merge(unique_more_than_one_contract,
                                       left_on='contractor',
                                       right_on='ВладелецИНН',
                                       how='inner')

united_df = contractors_df.merge(withdraw_docs,
                                 left_on='ИНН',
                                 right_on='contractor',
                                 how='inner')
united_etagi_df = contractors_etagi_df.merge(withdraw_docs,
                                             left_on='ИНН',
                                             right_on='contractor',
                                             how='inner')
aggregated_1 = (united_df
                .groupby(['contractor', 'Списание'],
                         as_index=False)
                .agg({'Дата': 'count'})
                .sort_values('Дата', ascending=False))

aggregated_etagi_1 = (united_etagi_df
                .groupby(['contractor', 'Списание'],
                         as_index=False)
                .agg({'Дата': 'count'})
                .sort_values('Дата', ascending=False))

diff_df = united_etagi_df[~united_etagi_df
.apply(lambda x: str(x.last_name).lower() in str(x.Наименование).lower(), axis=1)]

diff_df_2 = united_etagi_df[~united_etagi_df
.apply(lambda x: str(x.first_middle_name).lower() in str(x.Наименование).lower(), axis=1)]

diff_df_3 = diff_df[~diff_df
.apply(lambda x: str(x.first_middle_name).lower() in str(x.Наименование).lower(), axis=1)]
doubled_inn = aggregated_1[aggregated_1["Дата"] > 1].contractor
doubled_united = united_df[united_df['contractor'].isin(doubled_inn)]
doubled_etagi_inn = aggregated_etagi_1[aggregated_etagi_1["Дата"] > 1].contractor
doubled_etagi_united = united_etagi_df[united_etagi_df['contractor'].isin(doubled_etagi_inn)]

doubled_inn_number = doubled_united.groupby('Списание')\
                                   .agg({'Наименование': 'count'})\
                                   .shape[0]

incorrect_inn = withdraw_docs[
    (withdraw_docs.contractor.str.len() != 12) &
    (withdraw_docs.contractor.str.len() != 10)].shape[0]
not_numeric_inn = withdraw_docs[pd.to_numeric(withdraw_docs.contractor,
                                              errors='coerce').isna()].shape[0]
# amount of unrecognized withdraw documents
print(f'amount of doubled in DB by inn {doubled_inn_number}')
print(f'amount of incorrect inn in payment purpose {incorrect_inn}')
print(f'amount of not numeric inn {not_numeric_inn} in payment purpose (double checked)')
print('doubled contract exceptions {}'.format(doubled_contract.shape[0]))
total_exceptions_number = doubled_inn_number + incorrect_inn + doubled_contract.shape[0]
print('total exceptions amount {}'.format(total_exceptions_number))
print('the percentage of exceptions through the search{}'.format(total_exceptions_number /
                                                                 withdraw_docs.shape[0] * 100))

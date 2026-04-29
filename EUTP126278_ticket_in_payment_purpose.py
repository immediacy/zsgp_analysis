import pandas as pd
from sql_loads import drop_nan_columns

# define paths
ticket_path = (r"G:\Мой диск\Tasks"
               r"\EUTP-126278 для МТ-6699264 Обновить метод upsertNonCashPayments"
               r"\Заявки за 25-26 годы.xlsx")
payments_path = (r"G:\Мой диск\Tasks"
                 r"\EUTP-126278 для МТ-6699264 Обновить метод upsertNonCashPayments"
                 r"\Выгрузка в ОБП фильтр Покупатели.xlsx")

# read and prepare the data
payments_df = pd.read_excel(payments_path)
ticket_df = pd.read_excel(ticket_path)

payments_df = drop_nan_columns(payments_df)
ticket_df = drop_nan_columns(ticket_df)

# extract ticket from the comment
ticket_expression = r'\b\d{7,11}\b'
payments_df['ticket'] = payments_df.СчетНаОплатуКомментарий.str.findall(ticket_expression)
payments_df['ticket1'] = payments_df.ticket.str[0]
payments_df['ticket2'] = payments_df.ticket.str[1]
ticket_df = ticket_df.astype('str')
# check if there is a ticket number in the database
filtered_payments = payments_df[payments_df.ticket1.isin(ticket_df.НомерЗаявкиРИЭС)]
# extract a ticket form contract name
no_ticket_payments = payments_df[payments_df.ticket1.isna()]
no_ticket_payments.loc[:, 'contract_name_ticket'] = (no_ticket_payments
                                              .СчетНаОплатуДоговорКонтрагента
                                              .str
                                              .findall(ticket_expression))

no_ticket_payments.loc[:, 'ticket_contract'] = no_ticket_payments.contract_name_ticket.str[0]
filtered_contract_payments = no_ticket_payments[no_ticket_payments\
                                .ticket_contract.isin(ticket_df.НомерЗаявкиРИЭС)]


# extract ticket from the payment purpose
filtered_payments.loc[:, 'purp_ticket'] = filtered_payments.НазначениеПлатежа\
                                        .str.findall(ticket_expression)

payments_df['purp_ticket'] = payments_df.НазначениеПлатежа.str.findall(ticket_expression)
payments_df['purp_ticket'] = payments_df.purp_ticket.str[0]
purp_payment_df = payments_df[payments_df.purp_ticket.isin(ticket_df.НомерЗаявкиРИЭС)]

print(f'the amount of found tickets numbers in comment and contract is '
      f'{filtered_payments.shape[0] + filtered_contract_payments.shape[0]} of {payments_df.shape[0]}')
print(f'the amount of found tickets numbers in payment purpose is '
      f'{purp_payment_df.shape[0]} of {payments_df.shape[0]}')

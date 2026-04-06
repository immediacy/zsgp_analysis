import pandas as pd
log_path = (r"G:\Мой диск\Tasks\EUTP-89497 для МТ-6867575 Сопоставление поступлений со счетом на оплату"
            r"\Контроль метрик\ЗаменаКонтрагентовПодставлениеСчетов buh_tmn_2026-03-12.txt")
# prepare the file
with open(log_path) as _file:
    log = _file.read()

separated_log = log.split('\n')
# turn the file to the DF
log_df = pd.DataFrame(separated_log)
log_df.rename(columns={0: 'raw_str'}, inplace=True)
log_df = log_df.query('raw_str != "|0 сек"')
# devise the string to the sense items
log_df['splited'] = log_df.raw_str.str.split('|')
log_df.drop(8534, axis=0, inplace=True)
log_df['date_time'] = log_df.splited.map(lambda x: x[0])
log_df['action'] = log_df.splited.map(lambda x: x[2])
log_df['action'] = log_df.action.str.strip()
clean_log = log_df.iloc[:, 2: 4]
# let's find lines where we're looking for docs
in_doc_str = 'В документе'
in_doc_statement = clean_log[clean_log.action.str.contains(in_doc_str)]
not_in_doc_statement = clean_log[~clean_log.action.str.contains(in_doc_str)]
# let's find lines where we didn't find a ticket number in payment purpose in not_in_doc_statement df
no_ticket_str = "не найдено информации о заявке"
no_ticket_statement = not_in_doc_statement[not_in_doc_statement.action.str.contains(no_ticket_str)]
not_ticket_statement = not_in_doc_statement[~not_in_doc_statement.action.str.contains(no_ticket_str)]
# let's find replacement statements in @in_doc_statement df
replacement_str = 'заменен'
replacement_statement = in_doc_statement[in_doc_statement.action.str.contains(replacement_str)]
no_replacement_statement = in_doc_statement[~in_doc_statement.action.str.contains(replacement_str)]
# @no_replacement_statement contains only 'не совпадает с контрагентом' or 'не найдено записи в ЦОС'
no_replacement_statement_check = no_replacement_statement[
    (no_replacement_statement.action.str.contains('не совпадает с контрагентом') |
     no_replacement_statement.action.str.contains('не найдено записи в ЦОС'))
].shape[0]
print(f'except statements `не найдено записи в ЦОС` and `не совпадает с контрагентом`'
      f'no_replacement_statement df contains '
      f'{no_replacement_statement_check - no_replacement_statement.shape[0]} rows')
#
print(f'replacement counter {replacement_statement.shape[0]}')
print(f'no_replacement counter {no_replacement_statement.shape[0]}')
print(f'no_ticket_counter {no_ticket_statement.shape[0]}')
print(f'in_doc + no_ticket amount {no_ticket_statement.shape[0] + in_doc_statement.shape[0]}')

# metrics calculations
final_metric = replacement_statement.shape[0] / (in_doc_statement.shape[0] +
                                                 no_ticket_statement.shape[0])
print(f'final_metric {final_metric}')
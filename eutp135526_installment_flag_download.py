import pandas as pd

from sql_loads import drop_nan_columns

bi_path = r"G:\Мой диск\Tasks\EUTP-135526 Рассрочка для РТЗ\66.Все комиссии по сделкам_20-05-2026.xlsx"

bi_df = pd.read_excel(bi_path)

clean_df = drop_nan_columns(bi_df, 0)

clean_df.drop_duplicates()
bi_df_1 = clean_df.drop_duplicates(subset=['Номер сделки'])
bi_df_1.reset_index(drop=True, inplace=True)

new_columns = ['deal_number', 'ticket_number']
columns_dict = {bi_df_1.columns[i]: new_columns[i] for i in range(len(new_columns))}
bi_df_1.rename(columns=columns_dict, inplace=True)

rtz_dl = r"G:\Мой диск\Tasks\EUTP-135526 Рассрочка для РТЗ\Выгрузка - РТЗ (ссылка, дата, номер сделки, номер заявки) 2025 год.xlsx"
rtz_df = pd.read_excel(rtz_dl)
rtz_df_clean = drop_nan_columns(rtz_df)
merged = rtz_df_clean.merge(bi_df_1, how='outer', left_on='НомерСделки', right_on='deal_number')

one_s_lost = merged[merged.Ссылка.isna()]

available_installment = merged.dropna(axis=0, how='any')

print(f'available_installment.shape {available_installment.shape}')
print(f'one_s_lost shape {one_s_lost.shape}')
# print(f'bi_df.shape {bi_df.shape}')
print(f'bi_df_1.shape {bi_df_1.shape}')
print(f'rtz_df_clean.shape {rtz_df_clean.shape}')

print(merged.shape)

import pandas as pd
import os
# prepare the paths
my_path = r'G:\Мой диск\Tasks\EUTP-108187 для МТ-4934581 Доработка реквизитов департамент, направление и место учета затрат'
my_name = r'Соответствие статьи затрат.xlsx'
full_path = os.path.join(my_path, my_name)
data = pd.read_excel(full_path)

# let's prepare the data
def drop_nan_columns(dnc_data):
    # get a series with real table columns (not DF columns) and turn it to a dict
    real_columns = dnc_data.iloc[1, :].dropna()
    real_columns_dict = real_columns.to_dict()
    # drop table header: table columns names and time of execution statement
    dnc_data = dnc_data.drop([0, 1], axis=0).reset_index(drop=True)
    # drop NA columns from DF (appears after import xlsx to DF)
    dnc_data = dnc_data.dropna(axis=1).rename(columns=real_columns_dict)

    return dnc_data

data = drop_nan_columns(data)
# create a column with a flag: if True the expanses code is same
data['same_expanses_code'] = data.СтатьяДоговор == data.СтатьяЗатрат

# upload to xlsx
data.to_excel(
    os.path.join(
        my_path,
        'ready_comparison_of_expanses_codes.xlsx'
    ), index=False
)
diff_data = data[data.same_expanses_code == False]

diff_data.loc['r_contract'] = diff_data.СтатьяДоговор.str.lower().str.split('/')
diff_data['r_invoice'] = diff_data.СтатьяЗатрат.str.lower().str.split('/')
diff_data['contain'] = False
diff_data['contain'] = diff_data.apply(
    lambda x: bool(set(x['r_contract']) & set(x['r_invoice'])),
    axis=1
)
number_of_different = diff_data[diff_data.contain == False].shape[0]
print(f'diff_data {number_of_different / data.shape[0] * 100:.2f}%')


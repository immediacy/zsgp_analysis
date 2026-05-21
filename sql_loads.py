import pandas as pd

# let's prepare the data
def drop_nan_columns(dnc_data, header_row=1, na_rows='drop'):
    # get a series with real table columns (not DF columns) and turn it to a dict
    real_columns = dnc_data.iloc[header_row, :].dropna()
    real_columns_dict = real_columns.to_dict()
    # drop table header: table columns names and time of execution statement
    dnc_data = dnc_data.drop([i for i in range(header_row + 1)], axis=0).reset_index(drop=True)
    # drop NA rows in the values
    dnc_data = dnc_data.rename(columns=real_columns_dict)
    dnc_data = dnc_data.loc[:, list(real_columns_dict.values())]
    if na_rows == 'drop':
        temp_df = dnc_data[dnc_data.isna().any(axis=1)]
        if temp_df.shape[0] != 0:
            dnc_data = dnc_data.dropna(axis=0, how='any')
            print('dropped NaN rows:')
            print(temp_df.head())
    else:
        dnc_data = dnc_data.fillna(na_rows)
    # drop NA columns from DF (appears after import xlsx to DF)


    return dnc_data
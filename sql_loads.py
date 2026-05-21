import pandas as pd

# let's prepare the data
def drop_nan_columns(dnc_data, header_row=1):
    # get a series with real table columns (not DF columns) and turn it to a dict
    real_columns = dnc_data.iloc[header_row, :].dropna()
    real_columns_dict = real_columns.to_dict()
    # drop table header: table columns names and time of execution statement
    dnc_data = dnc_data.drop([i for i in range(header_row + 1)], axis=0).reset_index(drop=True)
    # drop NA columns from DF (appears after import xlsx to DF)
    dnc_data = dnc_data.dropna(axis=1).rename(columns=real_columns_dict)

    return dnc_data
import pandas as pd

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
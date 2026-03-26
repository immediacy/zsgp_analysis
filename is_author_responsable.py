import pandas as pd
import os

my_path = r'G:\Мой диск\Tasks\EUTP-108187 для МТ-4934581 Доработка реквизитов департамент, направление и место учета затрат'
my_name = r'Автор и ответственный по задаче на согласование счета.xlsx'
full_path = os.path.join(my_path, my_name)

data = pd.read_excel(full_path)
data = data.drop(['Unnamed: 1', 'Unnamed: 2'], axis=1)
data = data.drop([0, 1], axis=0).reset_index(drop=True)
data.rename(columns={'Unnamed: 0': 'responsible', 'Unnamed: 3': 'author'}, inplace=True)
data['r_split_r'] = data.responsible.str.lower().str.split(' ')
data['r_split'] = data.r_split_r.str[0]
data['a_split_r'] = data.author.str.lower().str.split(' ')
data['a_split'] = data.a_split_r.str[0]
data['a_split'] = data.a_split_r.str[0]

author_equal_responsible = data.query('r_split == a_split')
author_equal_responsible_num = author_equal_responsible.shape[0]
author_not_equal_responsible = data.query('r_split != a_split')
author_not_equal_responsible_num = author_not_equal_responsible.shape[0]

agg_authors = author_not_equal_responsible.groupby(['author', 'r_split']).agg({'a_split': 'count'})

print(f'tasks where authors is equal to responsible '
      f'{round(author_equal_responsible_num/data.shape[0] * 100, 2)}%')
data['contain'] = False
data['contain'] = data.apply(
    lambda x: bool(set(x['r_split_r']) & set(x['a_split_r'])),
    axis=1
)
author_equal_responsible_num_2 = data[data.contain].shape[0]

print(f'tasks where authors is equal to responsible second ' 
      f'{round(author_equal_responsible_num_2/data.shape[0] * 100, 2)}%')

print('hello world')
print(f'authors not equal to responsible ')
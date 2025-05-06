import time
import traceback

from g_connector import *
from access_db_service import *
import pandas as pd

# google sheet connection
client = connect_ss()
access_db_filename = r"C:\Users\User\Documents\zsgp_analysis\umto_db.accdb"
legal_form_list = ['ООО', 'АО', 'ИП']


# update_db_company_table takes
# a DataFrame - g_values with 14 columns n rows
# a DataFrame with 7 columns
def update_db_company_table(g_values, db_values, db_conn):
    g_values_c = g_values.copy()
    db_values_c = db_values.copy()
    company_columns = ["№",
                       "ИНН",
                       "КПП",
                       "Правовая форма",
                       "Наименование",
                       "Адрес склада",
                       "Заблокированна СКЗ"]
    utd_data = g_values_c.drop(g_values_c[g_values_c['ИНН']
                               .isin(db_values_c.inn.astype('int'))].index, axis=0)

    table_fields = ('id', 'inn', 'kpp', 'legal_form', 'company_name', 'legal_address', 'is_banned')
    try:
        for row_n in utd_data.index:

            temp_values = list(utd_data.loc[row_n, company_columns])
            temp_values[6] = bool(temp_values[6])  # is_banned column
            temp_values[1] = str(temp_values[1])  # inn column
            temp_values[2] = str(temp_values[2])  # kpp column
            temp_values[0] = int(temp_values[0])  # id column
            if (temp_values[3] == ''):
                for item in legal_form_list:
                    if item in temp_values[4]:
                        temp_values[4] = temp_values[4].replace(item, '').strip()
                        temp_values[3] = item
            insert_values(db_conn, 'company', table_fields, temp_values)
        return 0
    except:
        traceback.print_exc()
        return traceback.format_exc()


# company contacts table update function takes two dataframes and connection var
def update_db_company_contacts(g_values, db_values, db_conn):
    table_fields = ('company_contact_id', 'company_id', 'e_mail', 'tel_number', 'web_address')
    utd_data = g_values.drop(g_values[g_values['№']
                             .isin(db_values.company_id.astype('int'))].index, axis=0)
    max_contact_index = db_values.company_contact_id.max()
    company_columns = ['№', 'Телефон', 'Сайт', 'Адрес электронной почты']
    last_db_contact_index = int(db_values.company_contact_id.max())
    try:
        for row_n in utd_data.index:
            temp_values = list(utd_data.loc[row_n, company_columns])

            temp_values[3] = temp_values[3].strip()  # is_banned column
            # temp_values[1] = temp_values[1].strip()  # company_id column
            temp_values[2] = temp_values[2].strip()  # e_mail column
            temp_values[0] = int(temp_values[0])  # company_contact_id column
            last_db_contact_index += 1
            temp_values.insert(0, last_db_contact_index)
            insert_values(db_conn, 'company_contacts', table_fields, temp_values)

        return 0

    except:
        traceback.print_exc()
        return traceback.format_exc()


company_all_columns = ('id',
                       'inn',
                       'kpp',
                       'legal_form',
                       'company_name',
                       'legal_address',
                       'is_banned')
contacts_all_columns = ('company_contact_id',
                        'company_id',
                        'e_mail',
                        'tel_number',
                        'web_address')
# main loop
while True:

    if client != 0:
        # load company table data as a DataFrame
        g_company = read_values(client,
                                '1RtMz0_lYcl2wEY7JU5sZ882oLJJVWbKqTvfBY58d-zA',
                                'B1')
        g_company['Заблокированна СКЗ'] = g_company['Заблокированна СКЗ'].astype('bool')
        last_company_num = g_company['№'].max()
        last_inn = g_company[g_company['№'] == last_company_num].loc[:, 'ИНН'].iloc[0]

        # MS Access database connection
        db_connection = connect_to_access_db(access_db_filename)
        # load db company table

        company_tab_query = (f'SELECT {', '.join([i for i in company_all_columns])} ' +
                             f'FROM company')

        db_company_values = select_values(db_connection,
                                          'company',
                                          company_all_columns,
                                          company_tab_query)
        contacts_tab_query = (f'SELECT {", ".join([i for i in contacts_all_columns])} ' +
                              f'FROM company_contacts')
        db_company_contacts = select_values(db_connection,
                                            'company_contacts',
                                            contacts_all_columns,
                                            contacts_tab_query)
        last_db_company_num = db_company_values.id.max()
        last_db_contact_num = db_company_contacts.company_id.max()
        # db_company_values.inn = db_company_values.inn.astype('int')
        if last_company_num > last_db_company_num:
            update_db_company_table(g_company, db_company_values, db_connection)
        if last_company_num > last_db_contact_num:
            update_db_company_contacts(g_company, db_company_contacts, db_connection)

    time.sleep(10)

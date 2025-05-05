import pandas as pd
import pyodbc
import traceback

from pandas import DataFrame


def connect_to_access_db(db_file: str) -> pyodbc.Connection:
    """Устанавливает соединение с базой данных Access."""
    connection_string = (rf'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};'
                         rf'DBQ={db_file};')
    return pyodbc.connect(connection_string)


def execute_query(conn: pyodbc.Connection,
                  query: str,
                  q_tuple=None) -> DataFrame:
    """Выполняет SQL-запрос."""
    cursor = conn.cursor()
    if q_tuple is None:
        cursor.execute(query)
    else:
        cursor.execute(query, q_tuple)

    if 'SELECT' in query:
        temp_list = []
        temp_header = [column[0] for column in cursor.description]
        for i in cursor.fetchall():
            temp_list.append(list(i))

        return pd.DataFrame(data=temp_list, columns=temp_header)
    else:
        conn.commit()
        cursor.close()
        return 0


def close_connection(conn: pyodbc.Connection) -> None:
    """Закрывает соединение с базой данных."""
    conn.close()


# execute_query(conn, a_query, query_tuple)
# close_connection(conn)
def insert_values(conn, table_name, field_names, field_values, s_query=None):
    try:
        if s_query is None:
            a_query = (f'INSERT INTO {table_name} ({', '.join(field_names)}) ' +
                       f"VALUES ({', '.join(['?' for i in field_values])})")
            execute_query(conn, a_query, field_values)
        else:
            execute_query(conn, s_query)
        return 0
    except:
        print(field_values)
        traceback.print_exc()
        return traceback.format_exc()


def select_values(conn, table_name, field_names, s_query=None):
    try:
        if s_query is None:
            s_query = (f'SELECT ({', '.join(['?' for i in field_names])}) ' +
                       f'FROM {table_name}')
            result = execute_query(conn, s_query, field_names)
            return result
        else:
            result = execute_query(conn, s_query)
            return result
    except:
        traceback.print_exc()
        return traceback.format_exc()

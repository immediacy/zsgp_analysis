import json
import os


def fix_and_parse(raw_string):
    # 1. Защищаем внутренние кавычки компании (с тремя слэшами \\\\\")
    step = raw_string.replace('\\\\\\"', 'TEMP_INNER_QUOTE')
    # 2. Удаляем строковые символы переноса
    step = step.replace('\\r\\n', '')
    # 3. Превращаем обычные экранированные кавычки ключей в валидные кавычки JSON
    step = step.replace('\\"', '"')
    # 4. Возвращаем внутренние кавычки компании с ОДНИМ корректным слэшем (\")
    step = step.replace('TEMP_INNER_QUOTE', '\\"')
    # 5. Безопасно парсим чистую строку в словарь
    return json.loads(step)

def find_ticket(req_path):


    with open(req_path) as txt_file:
        req_cont = txt_file.read()

    req_list = req_cont.split('|')
    req_list_1 = []
    for i in req_list:
        if 'upsertNonCashPayments' in i:
            req_list_1.append(i)


    temp_upsert_list = []
    for j in req_list_1:
        temp_j = j.split('"body": ')
        temp_j[1] = temp_j[1].strip('}')
        temp_j[1] = temp_j[1].strip('"')
        temp_j[1] = temp_j[1].strip('"\n')
        temp_upsert_list.append(temp_j[1])

    # Обрабатываем весь ваш список строк
    dicts_list = [fix_and_parse(row) for row in temp_upsert_list]
    # print(req_path)
    for l in dicts_list:
        if 'ticketId' in l.keys():
            print(l['statNumber'])
        if 'ЛН00-000201' in l['statNumber']:
            print(l)

folder_path =  r"\\s2d\Log\buh_tmn"
file_names = os.listdir(folder_path)
search_key = 'Запросы_Космос_buh_tmn_2026-'
found_names = []
for f in file_names:
    if search_key in f and '.txt' in f:
        found_names.append(f)

for file_name in found_names:
    path_template = os.path.join(folder_path, file_name)
    find_ticket(path_template)


import xlrd, re

xl_path = r"C:\Users\User\Downloads\Копия КО - финал_СО583-2-391(1).xls"


def ko_extractor(x_path) -> list:
    ko_list = []
    ss = xlrd.open_workbook(x_path)
    sheet = ss.sheet_by_index(0)
    for rw in range(sheet.nrows):
        ko_list.append(sheet.row(rw))

    return ko_list


def first_table_row_flag(ko_list, t_row):
    if 'наименование' in ko_list[t_row][1].value.lower():

        return True, t_row
    else:
        for i in range(t_row, t_row + 10):
            if 'наименование' in ko_list[i][1].value.lower():
                return True, t_row


def ko_parser(k_path):
    k_list = ko_extractor(k_path)
    k_list_len = len(k_list)
    ko_dict = {}
    com_dict = {'goods': {'g_name': {}, 'volume': {}, 'unit': {}}}
    pp_row = 7
    ko_parser_dict = {'ko_date': 'дата составления',
                      'object': 'объект:',
                      'order': 'заявка:', }
    table_parser_dict = {'inn': 'ИНН',
                         'supplier_name': 'Поставщик',
                         'place': 'Местонахождение',
                         'table_start': 'Наименование и тех. хар-ка',
                         'table_end': 'Итого, сумма в рублях с НДС',
                         'delivery_place': 'базис поставки',
                         'delivery_flag': 'Доставка до Базиса',
                         'warranty': 'Гарантийный срок',
                         'zsgp_contract': 'Работа по  форме договора ЗГСП',
                         'edo': 'Подписание договора в ЭДО ',
                         'payment_conditions': 'Условия оплаты'}

    # order and project parsing part
    if (ko_parser_dict['ko_date'] in k_list[3][1].value.lower() and
            ko_parser_dict['object'] in k_list[4][1].value.lower() and
            ko_parser_dict['order'] in k_list[5][1].value.lower()):
        ko_dict['ko_date'] = k_list[3][1].value.split(':')[1].strip()
        ko_dict['object'] = k_list[4][1].value.split(':')[1].strip()
        ko_dict['order'] = k_list[5][1].value.split(':')[1].replace(' ', '').replace('№', '')
    else:
        for na_row, a_row in enumerate(k_list):
            if a_row[0].value == 'п/п':
                pp_row = na_row
                break
            for cell in a_row:
                if cell.value != '':

                    for key, val in table_parser_dict.items():
                        if val in str(cell.value).lower():
                            if ':' in str(cell):
                                ko_dict[key] = cell.value.split(':')[1]
                                break
                            else:
                                ko_dict[key] = cell.value
                                break
    # header parsing part

    inn_row = pp_row + 1
    inn_col = []
    inn_exp = r'[\d]{10,12}'
    if 'инн' in k_list[inn_row][1].value.lower():
        for nb_cell, b_cell in enumerate(k_list[inn_row]):
            inn_match = re.search(inn_exp, str(b_cell.value))
            if inn_match is not None:
                inn_col.append(nb_cell)
                com_dict[nb_cell] = {
                    'inn': inn_match[0],
                    'name': k_list[inn_row - 1][nb_cell].value,
                    'place': k_list[inn_row + 1][nb_cell].value,
                    'goods': {'price': {}, 'sum': {}, 'made': {}, 'ost': {}}
                }

    else:

        for n_row in range(pp_row, k_list_len):

            for n_cell, b_cell in enumerate(k_list[n_row]):
                if 'инн' in str(b_cell.value).lower():
                    inn_row = n_row
                    inn_match = re.search(inn_exp, str(b_cell.value))
                    if inn_match is not None:
                        inn_col.append(n_cell)
                        com_dict[n_cell] = {
                            'inn': inn_match[0],
                            'name': k_list[inn_row - 1][n_cell].value,
                            'place': k_list[inn_row + 1][n_cell].value,
                            'goods': {'price': {}, 'sum': {}, 'made': {}, 'ost': {}}
                        }
    table_row = inn_row + 2

    first_row_flag, table_row = first_table_row_flag(k_list, table_row)

    end_of_table = table_row + 2
    if first_row_flag:
        for nc_row in range(table_row + 1, k_list_len):
            if 'итого' in k_list[nc_row][1].value.lower():
                end_of_table = nc_row
                break
            com_dict['goods']['g_name'][nc_row] = k_list[nc_row][1].value
            com_dict['goods']['unit'][nc_row] = k_list[nc_row][2].value
            com_dict['goods']['volume'][nc_row] = k_list[nc_row][3].value

            for nc_cell, c_cell in enumerate(k_list[nc_row]):
                # the selector below make a choice which cell to write down
                if nc_cell in com_dict.keys():
                    if c_cell.value:
                        com_dict[nc_cell]['goods']['price'][nc_row] = c_cell.value
                    continue

                elif nc_cell - 1 in com_dict.keys():
                    if c_cell.value:
                        com_dict[nc_cell - 1]['goods']['sum'][nc_row] = c_cell.value
                    continue

                elif nc_cell - 2 in com_dict.keys():
                    if c_cell.value:
                        com_dict[nc_cell - 2]['goods']['made'][nc_row] = c_cell.value
                    continue

                elif nc_cell - 3 in com_dict.keys():
                    if c_cell.value:
                        com_dict[nc_cell - 3]['goods']['ost'][nc_row] = c_cell.value
                    continue
        break_flag = False
        for nd_row in range(end_of_table, k_list_len):
            if break_flag:
                break
            for td_key, td_val in table_parser_dict.items():
                temp_first_val = k_list[nd_row][1].value.lower()
                if td_val.lower() in temp_first_val:
                    for col_n in inn_col:
                        com_dict[col_n][td_key] = k_list[nd_row][col_n].value
                elif 'предлагается' in temp_first_val:
                    for c_num in inn_col:
                        if com_dict[c_num]['name'].lower() in temp_first_val:
                            com_dict[c_num]['winner'] = True
                        else:
                            com_dict[c_num]['winner'] = False
                    break_flag = True
                    break

    else:
        print('table_row is not defined')


    return ko_dict


ko_parser(xl_path)
print(__name__)

import fitz  # pip install pymupdf
import re, os

all_legal_form_exp = (r'(?i)\b(ООО|АО|ЗАО|ОАО|ПАО|ИП|НКО'
                       r'|ТОО|ГК|СОЮЗ|ФОНД|ТД|ТФ|'
                       r'Акционерное общество|Компания|Общество с ограниченной ответственностью|'
                       r'Публичное акционерное общество|Индивидуальный предприниматель|'
                       r'Группа компаний)')
short_legal_form_exp = (r'(?i)\b(ООО|АО|ЗАО|ОАО|ПАО|ИП|НКО'
                        r'|ТОО|ГК|СОЮЗ|ФОНД|ТД|ТФ|)')
full_legal_form_exp = (r'(?i)\b(Акционерное общество|Компания|Общество с ограниченной ответственностью|'
                       r'Публичное акционерное общество|Индивидуальный предприниматель|Группа компаний)')

only_name_exp = r'[\s«"„“”\'‹›„“”‘’]*([\w\s\-.&«»"“”‘’„]+)\b'
only_name_exp_2 = r'\s*["«„“”\'»][^"«„“”\'»]+["»“”\'»]'
company_name_exp = all_legal_form_exp + only_name_exp
company_name_exp_2 = all_legal_form_exp + only_name_exp_2

invoice_exp = (r'(?i)Счет(?: на оплату)? №\s*([A-Za-zА-Яа-я0-9\-\/]+)\s+'
               r'от\s+(\d{1,2}\s+[а-яА-Я]+(?:\s+\d{4})?)\s*г?\.?')

inn_exp = r'(?i)\bИНН\b[:\s]*\d{10}\b'
inn_exp_ip = r'(?i)\bИНН\b[:\s]*\d{12}\b'
kpp_exp = r'(?i)\bКПП\b[:\s]*\d{9}\b'
address_exp = [(r'(?i)(г\.|пос\.|пгт\.|с\.|город|р\.п\.)\s+[А-ЯЁа-яё0-9\s\-А-ЯЁа-яё0-9\s\],\s+'
                r'(ул\.|пр\.|пер\.|шоссе|наб\.|пл\.|бульвар|проспект|дорога|тупик|проезд|улица)'
                r'\s+[А-ЯЁа-яё0-9\s\-/]+'),
               (r'(?i)[А-ЯЁа-яё0-9\s\-А-ЯЁа-яё0-9\s\]+\s(г\.|пос\.|пгт\.|с\.|город|р\.п\.),\s+'
                r'(ул\.|пр\.|пер\.|шоссе|наб\.|пл\.|бульвар|проспект|дорога|тупик|проезд|улица)'
                r'\s+[А-ЯЁа-яё0-9\s\-/]+')]
tel_exp = r'(?i)(тел.*)?((8|\+7)[\- ]?)?(\(?\d{3}\)?[\- ])[\d\- ]{7,10}'
web_exp = r'(?<!\S)(\w*-?\w*\.?\w*-?\w*\.?\w*-?\w*[a-zA-Z]\.[\w-]{2,3})\b'
mail_exp = r'(?<!\S)([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)'

legal_form_dict = {'Общество с ограниченной ответственностью':'ООО',
                   'Акционерное общество': 'АО',
                   'Публичное акционерное общество': 'ПАО',
                   'Индивидуальный предприниматель': 'ИП',
                   'Группа компаний': 'ГК'}


# the func for invoice_matter_extractor
# in takes match as essentials, dict and key to check if it is in the dict
# if not, add in the essentials in the list
# and if contrary add the key and put the list with essentials
def dict_adder(key, _dict, reg_match):
    if key in _dict.keys():
        if type(reg_match) is str:
            _dict[key].append(reg_match)
        else:
            _dict[key].append(reg_match[0])
    else:
        if type(reg_match) is str:
            _dict[key] = [reg_match]
        else:
            _dict[key] = [reg_match[0]]


# the function intended to delete zsgp company information from the list
# once the essentials has been extracted from the page
def company_name_check(name_exp, item, r_dict):
    black_list = ['запсибгазпром',
                  'банк',
                  'ГК РФ',
                  'ГКРФ']

    company_name = re.search(name_exp, item)
    if company_name is None:
        return None

    elif 'company_name' in r_dict.keys():
        if (r_dict['company_name'][0] in company_name[0] or
                company_name[0] in r_dict['company_name'][0]):
            return None

    for b in black_list:
        if b in company_name[0].lower():
            return None

    if len(company_name[0]) > 6:
        return company_name

    else:
        return None


def address_check(address_exp_c, item, r_dict: dict, reg_num=0):
    address = re.search(address_exp_c[reg_num], item)
    zsgp_address = {'city': 'Тюмень',
                    'street_type': 'тракт',
                    'street': 'Велижанский',
                    'building_num': '11'}
    black_list = ['Можайск',
                  'Ильинское',
                  'Шексн']
    counter = 0
    cnt_bit = 1/len(zsgp_address)
    if address is None:
        return None
    if 'address' in r_dict.keys():
        if len(r_dict['address'][0]) == len(address[0]):
            return None
        elif len(r_dict['address'][0]) > len(address[0]):
            return None
        else:
            pass
    for black in black_list:
        if black in address[0]:
            return None
    for key in zsgp_address.values():
        if key in address[0]:
            counter += cnt_bit
        elif len(address[0]) > 10:
            return address
        elif reg_num + 1 < len(address_exp_c):
            return address_check(address_exp_c, item, r_dict, reg_num + 1)
        else:
            return address
    if counter > 0.7:
        return None


def kpp_check(kpp_exp_, item, r_dict: dict):
    zsgp_kpp = '720301001'
    kpp_ = re.search(kpp_exp_, item)
    validator_kpp = r'[\d]+'
    if kpp_ is None or zsgp_kpp in kpp_[0]:
        return None
    elif 'kpp' in r_dict.keys():
        if r_dict['kpp'][0] in kpp_[0]:
            return None
    else:
        kpp_clean = re.search(validator_kpp, kpp_[0])
        return kpp_clean


def inn_check(inn_exp_, item, r_dict: dict):
    zsgp_inn = '7202083210'
    inn_ = re.search(inn_exp_, item)
    validator_inn = r'[\d]+'

    if inn_ is None or zsgp_inn in inn_[0]:
        return None
    elif 'inn' in r_dict.keys():
        if r_dict['inn'][0] in inn_[0]:
            return None
    else:
        inn_clean = re.search(validator_inn, inn_[0])
        return inn_clean


def tel_check(tel_exp_, item, r_dict: dict):
    zsgp_num = ['3452540541']
    tel_ = re.search(tel_exp_, item)
    sharp_list = 'тел.,: ()-'
    validator_reg = r'[\d\-\+\(\)\s]+'

    zsgp_num_cnt = 1 / len(zsgp_num)
    tel_counter = 0
    if tel_ is None:
        return None
    else:
        tel_res = re.search(validator_reg, tel_[0])
        tel_clean = tel_[0].strip()
        if len(tel_clean) < 10:
            return None
        for l in sharp_list:
            tel_clean = tel_clean.replace(l, '')
        for t_number in zsgp_num:
            if t_number in tel_clean:
                return None
            else:
                tel_counter += zsgp_num_cnt

        return tel_res[0].strip()


# the function intended to extract the matter company information
def invoice_matter_extractor(invoice_path):
    res_dict = {}
    doc = fitz.open(invoice_path)
    text = []
    for page in doc:
        [text.append(i) for i in page.get_text().split('\n')]
    doc.close()

    for item in text:

        company_re = company_name_check(company_name_exp, item, res_dict)

        inn_re = inn_check(inn_exp, item, res_dict)
        inn_re_ip = inn_check(inn_exp_ip, item, res_dict)
        kpp_re = kpp_check(kpp_exp, item, res_dict)
        address_re = address_check(address_exp, item, res_dict)
        tel_re = tel_check(tel_exp, item, res_dict)
        web_re = re.search(web_exp, item)
        mail_re = re.search(mail_exp, item)
        invoice_num = re.search(invoice_exp, item)

        if company_re is not None:
            dict_adder('company_name', res_dict, company_re)

        if inn_re is not None:
            dict_adder('inn', res_dict, inn_re)

        if inn_re_ip is not None:
            dict_adder('inn_ip', res_dict, inn_re_ip)

        if kpp_re is not None:
            dict_adder('kpp', res_dict, kpp_re)

        if address_re is not None:
            dict_adder('address', res_dict, address_re)

        if tel_re is not None:
            dict_adder('tel', res_dict, tel_re)

        if web_re is not None:
            dict_adder('web_address', res_dict, web_re)

        if mail_re is not None:
            dict_adder('e_mail', res_dict, mail_re)

        if invoice_num is not None:
            dict_adder('invoice', res_dict, invoice_num)

    return res_dict


def parencies_cleaner(a_str):
    cleaner_str = r'«"„“”\'‹›„“”‘’»'
    temp_item = a_str
    for i in cleaner_str:
        if i in a_str:
            temp_item = temp_item.replace(i, '')
    return temp_item


def add_name_to_dict(short_l, full_l, name_c):
    res_dict = {'legal': '',
                'name': '',
                'fill_flag': 0}
    if short_l is not None:
        short_l = short_l[0]
        res_dict['legal'] = short_l
        res_dict['name'] = name_c.replace(short_l, '').strip()
        res_dict['fill_flag'] = bool(short_l) + bool(res_dict['name'])
    elif (full_l is not None) and (full_l in legal_form_dict.keys()):
        full_l = full_l[0]
        res_dict['legal'] = legal_form_dict[full_l]
        res_dict['name'] = name_c.replace(short_l, '').srip()
        res_dict['fill_flag'] = bool(res_dict['legal']) + bool(res_dict['name'])
    elif full_l not in legal_form_dict.keys():
        print('not in legal dict')
    else:
        print('short and full re results probably empty')

    return res_dict


def name_chooser(r_dict, comp_exp):
    # put the company list from the result dict
    com_list = r_dict['company_name']
    # define global var for the function result
    res_com = {'legal': '',
               'name': '',
               'fill_flag': 0}
    # iterate the company list and find the best option for company name
    for i in com_list:

        temp_com = re.search(comp_exp, i)
        short_legal = re.search(short_legal_form_exp, i)
        full_legal = re.search(full_legal_form_exp, i)
        if temp_com is not None:
            clean_name = parencies_cleaner(temp_com[0])
            clean_list = clean_name.split()
            clean_cnt = 1/len(clean_list)
            name_cnt = 0
            for j in clean_list:
                if j in short_legal_form_exp:
                    name_cnt += clean_cnt
        else:
            name_cnt = 1

        if temp_com is not None and name_cnt < 0.95:

            if res_com['fill_flag'] == 0:
                temp_res_com = add_name_to_dict(short_legal, full_legal, clean_name)
                if temp_res_com['fill_flag'] > 0:
                    res_com = temp_res_com
            elif res_com['fill_flag'] == 1:
                temp_res_com = add_name_to_dict(short_legal, full_legal, clean_name)
                if temp_res_com['fill_flag'] > 1:
                    res_com = temp_res_com
            else:
                temp_res_com = add_name_to_dict(short_legal, full_legal, clean_name)
                s_name_condition = (res_com['legal'] != temp_res_com['legal'] and
                                    res_com['name'] != temp_res_com['name'])
                if temp_res_com['fill_flag'] == 2 and s_name_condition:
                    print('second possible company name found')
        else:
            clean_name = parencies_cleaner(i)
            if res_com['fill_flag'] == 0:
                temp_res_com = add_name_to_dict(short_legal, full_legal, clean_name)
                if temp_res_com['fill_flag'] > 0:
                    res_com = temp_res_com
            elif res_com['fill_flag'] == 1:
                temp_res_com = add_name_to_dict(short_legal, full_legal, clean_name)
                if temp_res_com['fill_flag'] > 1:
                    res_com = temp_res_com
            else:
                temp_res_com = add_name_to_dict(short_legal, full_legal, clean_name)
                s_name_condition = (res_com['legal'] != temp_res_com['legal'] and
                                    res_com['name'] != temp_res_com['name'])
                if temp_res_com['fill_flag'] == 2 and s_name_condition:
                    print('second possible company name found')
    return [res_com['legal'], res_com['name']]

def address_chooser(r_dict):
    address_list = r_dict['address']
    max_len = len(address_list[0])
    res_addres = address_list[0]
    for i in address_list:
        if max_len < len(i):
            max_len = len(i)
    return res_addres

ex_path = r"C:\Users\User\Downloads\drive-download-20250506T041516Z-001"

path_list = os.listdir(ex_path)
res_list = []
for file in path_list:
    temp_item = invoice_matter_extractor(os.path.join(ex_path, file))
    if 'company_name' in temp_item.keys():
        temp_item['company_name'] = name_chooser(temp_item, company_name_exp_2)
    if 'address' in temp_item.keys():
        temp_item['address'] = address_chooser(temp_item)
    res_list.append(temp_item)

cnt = 0
for n, i in enumerate(res_list):
    if 'company_name' not in i.keys():
        cnt += 1
        print(n)
print(cnt)

print('a')

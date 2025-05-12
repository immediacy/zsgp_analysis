import fitz  # pip install pymupdf
import re, os

legal_form_exp = (r'(?i)\b(ООО|АО|ЗАО|ОАО|ПАО|ИП|НКО|Общество с ограниченной ответственностью'
                    r'|ТОО|ГК|СОЮЗ|ФОНД|ТД|ТФ|Акционерное общество|Компания|'
                    r'Публичное акционерное общество|Индивидуальный предприниматель)')
company_name_exp = (legal_form_exp +
                    r'[\s«"„“”\'‹›„“”‘’]*([\w\s\-.&«»"“”‘’„]+)\b')
company_name_exp_2 = legal_form_exp + r'\s*["«„“”\'»][^"«„“”\'»]+["»“”\'»]'
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


# the func for invoice_matter_extractor
# in takes match as essentials, dict and key to check if it is in the dict
# if not, it add in the essentials in the list
# and if contrary it add the key and put the list with essentials
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


def address_check(address_exp, item, r_dict: dict, reg_num=0):
    address = re.search(address_exp[reg_num], item)
    zsgp_address = {'city': 'Тюмень',
                    'street_type': 'тракт',
                    'street': 'Велижанский',
                    'building_num': '11'}
    counter = 0
    if address is None:
        return None
    if 'address' in r_dict.keys():
        if len(r_dict['address'][0]) == len(address[0]):
            return None
        elif len(r_dict['address'][0]) > len(address[0]):
            return None
        else:
            pass

    for key in zsgp_address.keys():
        if zsgp_address[key] in address[0]:
            counter += 0.25
        elif len(address[0]) > 10:
            return address
        elif reg_num < len(address_exp):
            return address_check(address_exp[reg_num + 1], item, r_dict)
    if counter == 1:
        return None

def kpp_check(kpp_exp_, item, r_dict: dict):
    zsgp_kpp = '720301001'
    kpp_ = re.search(kpp_exp_, item)

    if kpp_ is None or zsgp_kpp in kpp_[0]:
        return None
    elif 'kpp' in r_dict.keys():
        if r_dict['kpp'][0] in kpp_[0]:
            return None
    else:
        if 'КПП ' in kpp_[0]:
            return kpp_[0].replace('КПП ', '')
        elif 'КПП' in kpp_[0]:
            return kpp_[0].replace('КПП', '')
        else:
            return kpp_

def inn_check(inn_exp_, item, r_dict: dict):
    zsgp_inn = '7202083210'
    inn_ = re.search(inn_exp_, item)
    validator_inn = r'[\d]'

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

    return res_dict


ex_path = r"C:\Users\User\Downloads\drive-download-20250506T041516Z-001"
path_list = os.listdir(ex_path)
res_list = []
for file in path_list:
    res_list.append(invoice_matter_extractor(os.path.join(ex_path, file)))

cnt = 0
for n, i in enumerate(res_list):
    if 'company_name' not in i.keys():
        cnt += 1
        print(n)
print(cnt)

print('a')

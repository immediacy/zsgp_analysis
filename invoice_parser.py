import fitz  # pip install pymupdf
import re, os


company_name_exp = (r'(?i)\b(ООО|АО|ЗАО|ОАО|ПАО|ИП|НКО|Общество с ограниченной ответственностью'
                    r'|ТОО|ГК|СОЮЗ|ФОНД|ТД|ТФ|Акционерное общество|Компания|'
                    r'Публичное акционерное общество|Индивидуальный предприниматель)\b'
                    r'[\s«"„“”\'‹›„“”‘’]*([\w\s\-.&()«»"“”‘’„]+)')
inn_exp = r'(?i)\bИНН\b[:\s]*\d{10}\b'
inn_exp_ip = r'(?i)\bИНН\b[:\s]*\d{12}\b'
kpp_exp = r'(?i)\bКПП\b[:\s]*\d{9}\b'
address_exp = (r'\b(ул\.|пр\.|пер\.|шоссе|наб\.|пл\.|бульвар'
                     r'|проспект|дорога|тупик)\s+[А-ЯЁа-яё0-9\s\-]+\s*\d+\b')
tel_exp = r'^((8|\+7)[\- ]?)?(\(?\d{3}\)?[\- ]?)?[\d\- ]{7,10}$'
web_exp = r'\b(?:www\.)?[\w-]+(\.[\w-]+)+\b'
mail_exp = r'([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9_-]+)'

def dict_adder(key, _dict, reg_match):

    if key in _dict.keys():
        _dict[key].append(reg_match[0])
    else:
        _dict[key] = [reg_match[0]]

def invoice_matter_extractor(invoice_path):

    res_dict = {}
    doc = fitz.open(invoice_path)
    text = []
    for page in doc:
        [text.append(i) for i in page.get_text().split('\n')]
    doc.close()

    for item in text:

        company_re = re.search(company_name_exp, item)
        inn_re = re.search(inn_exp, item)
        inn_re_ip = re.search(inn_exp_ip, item)
        kpp_re = re.search(kpp_exp, item)
        address_re = re.search(address_exp, item)
        tel_re = re.search(tel_exp, item)
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

print('a')

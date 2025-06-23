import fitz  # pip install pymupdf
import re, os

supplier_list = ['Поставщик']
buyer_list = ['Заказчик', 'Покупатель']
receiver_list = ['Получатель']
base = ['']
table_end = ['Итого', 'Всего']

def name_cleaner(given_name, legal):
    res_name = given_name

    return res_name
def main_extractor(invoice_folder_path):
    invoice_dict = {}
    folder_files = os.listdir(invoice_folder_path)
    company_details_dict = {}
    for file in folder_files:
        invoice_path = os.path.join(invoice_folder_path, file)
        doc = fitz.open(invoice_path)
        text = ''
        for page in doc:
            text += page.get_text()
        doc.close()
        invoice_dict[file] = text

        # company details extraction
        details_start = text.find(supplier_list[0]) + len(supplier_list[0]) + 1
        details_stop = details_start + 300
        buyer_label = ''
        for buyer in buyer_list:
            if buyer in text:
                details_stop = text.find(buyer)
                buyer_label = buyer
        # company details text to list
        com_details = text[details_start:details_stop].split(',')
        # company details extraction loop
        for item in com_details:
            inn_exp = r'[\d]{10,12}'
            inn_match = re.search(inn_exp, item)
            if inn_match is not None:
                company_details_dict[file] = {'inn': inn_match[0]}

            kpp_exp = r'[\d]{9}'
            kpp_match = re.search(kpp_exp, item)
            if kpp_match is not None:
                if file in company_details_dict.keys():
                    company_details_dict[file]['kpp'] = kpp_match[0]
                else:
                    company_details_dict[file] = {'kpp': kpp_match[0]}

            legal_form_exp = (r'(?i)\b(ООО|АО|ЗАО|ОАО|ПАО|ИП|НКО'
                              r'|ТОО|ГК|СОЮЗ|ФОНД|ТД|ТФ|'
                              r'Акционерное общество|Компания|Общество с ограниченной ответственностью|'
                              r'Публичное акционерное общество|Индивидуальный предприниматель|'
                              r'Группа компаний)')
            legal_match = re.search(legal_form_exp, item)
            if legal_match is not None:
                temp_name = item[item.find(legal_match[0]) + len(legal_match[0]):]
                if file in company_details_dict.keys():

                    company_details_dict[file]['name'] = temp_name
                    company_details_dict[file]['legal'] = legal_match[0]
                else:
                    company_details_dict[file]['name'] = temp_name
                    company_details_dict[file] = {'name': legal_match[0]}
            validator_reg = r'[\d\-\+\(\)\s]+'
            tel_exp = r'(?i)(тел.*)?((8|\+7)[\- ]?)?(\(?\d{3}\)?[\- ])[\d\- ]{7,10}'
            tel_match = re.search(tel_exp, item)
            if tel_match is not None:
                temp_tel = re.search(validator_reg, tel_match[0])
                if file in company_details_dict.keys():
                    company_details_dict[file]['tel'] = temp_tel[0]
                else:
                    company_details_dict[file] = {'tel': temp_tel[0]}

            invoice_exp = (r'(?i)Счет(?: на оплату)? №\s*([A-Za-zА-Яа-я0-9\-\/]+)\s+'
                           r'от\s+(\d{1,2}\s+[а-яА-Я]+(?:\s+\d{4})?)\s*г?\.?')
            invoice_match = re.search(invoice_exp, item)
            if invoice_match is not None:
                if file in company_details_dict.keys():
                    company_details_dict[file]['invoice'] = invoice_match[0]
                else:
                    company_details_dict[file] = {'invoice': invoice_match[0]}
        # goods items extraction
        body_start = details_stop + len(buyer_label)
        body_text = text[body_start:]
        body_end = len(body_text)
        for marker in table_end:
            if marker in body_text:
                body_end = body_text.find(marker)
        body_text = body_text[:body_end]
        body_list = body_text.split('\n')
        header_list = [['№'], ['Наименование'], ['Кол-во'], ['Цена'], ['Сумма']]
        # for cell in body_list:


    return invoice_dict


inv_dict = main_extractor(r'C:\Users\User\Downloads\drive-download-20250506T041516Z-001')


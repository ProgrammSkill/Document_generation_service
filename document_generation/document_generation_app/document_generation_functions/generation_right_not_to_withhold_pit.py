import os
from datetime import datetime
from openpyxl import *
import pytz
from document_generation_app.document_generation_functions.api import CompanyAPI, IndividualAPI
from document_generation_app.document_generation_functions.functions import Date_conversion, Get_path_file

path_file = Get_path_file()

def Generation_generation_right_not_to_withhold_pit(request):
    company = CompanyAPI()
    organization = company["organizationalForm"] + ' "' + company["name"] + '"'
    list_organization = list(organization.upper())
    phone = company["contactInfo"]["phone"].upper()
    list_phone = list(phone)
    inn = company['inn']
    list_inn = list(inn)
    kpp = company['kpp']
    ogrn = company['ogrn']
    list_ogrn = list(ogrn)
    paymentAccount = company['bank']['paymentAccount'].upper()
    correspondentAccount = company['bank']['correspondentAccount'].upper()
    legalAddress = company['legalAddress']["name"].upper()
    list_legalAddress = list(legalAddress)
    ActualAddreses = company['ActualAddreses'][0]["name"].upper()
    CEO = 'Сталюкова Екатерина Александровна'

    individual = 'Гайназаров Кайратбек'.upper()
    # passport_series = 'AC'.upper()
    # passport_number = '4348554'.upper()

    code = request.POST.get('code')
    list_code = list(code)

    path_file_doc = 'document_generation_app/document_templates/right_not_to_withhold_pit.xlsx'
    doc = load_workbook(path_file_doc)
    sheet = doc.active

    list_columns = ['X', 'AA', 'AD', 'AG', 'AJ',
                    'AM', 'AP', 'AS', 'AV', 'AY', 'BB', 'BH']

    row = 2
    index = 0
    for symbol in list_inn:
        for col in range(index, len(list_columns)):
            cell = list_columns[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            break

    list_columns = ['X', 'AA', 'AD', 'AG', 'AJ',
                    'AM', 'AP', 'AS', 'AV']
    row = 4
    index = 0
    for symbol in list_ogrn:
        for col in range(index, len(list_columns)):
            cell = list_columns[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            break

    # Input code
    sheet["CF8"] = list_code[0]
    sheet["CJ8"] = list_code[1]
    sheet["CO8"] = list_code[2]
    sheet["CT8"] = list_code[3]

    list_columns = ['B', 'C', 'E', 'F', 'G', 'I', 'L', 'N', 'P', 'S', 'U', 'W', 'Z', 'AC', 'AF', 'AI', 'AL',
                    'AO', 'AR', 'AU', 'AX', 'AY', 'BA', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BE', 'BJ', 'BK', 'BL',
                    'BM', 'BO', 'BP', 'BR', 'BS', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CE']
    row = 11
    index = 0
    for symbol in list_organization:
        for col in range(index, len(list_columns)):
            cell = list_columns[index] + str(row)
            if list_columns[index] == 'CE':
                sheet[f'{cell}'] = symbol
                row += 2
                index = 0
                break
            sheet[f'{cell}'] = symbol
            index += 1
            break

    now_date = datetime.now(pytz.timezone('UTC'))
    now_date = Date_conversion(now_date.strftime('%d-%m-%Y'))
    list_now_date = list(now_date)

    list_columns = ['U', 'W', 'Z', 'AC', 'AF',
                    'AI', 'AL', 'AO', 'AR', 'AU']
    row = 46
    index = 0
    for symbol in list_now_date:
        for col in range(index, len(list_columns)):
            cell = list_columns[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            break

    global path_file
    path = path_file
    if os.path.exists(path + '/' + 'right_not_to_withhold_pit.xlsx') == False:
        doc.save(path + '/' + 'right_not_to_withhold_pit.xlsx')
    else:
        i = 1
        while True:
            if os.path.exists(path_file + '/' + f'right_not_to_withhold_pit{i}.xlsx') == False:
                path = path_file + '/' + f'right_not_to_withhold_pit{i}.xlsx'
                doc.save(path)
                break
            i += 1
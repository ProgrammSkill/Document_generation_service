import os
from datetime import datetime
from openpyxl import *
import pytz
from document_generation_app.document_generation_functions.api import CompanyAPI, IndividualAPI
from document_generation_app.document_generation_functions.functions import Date_conversion, Get_path_file
from django.http import FileResponse, HttpResponse, StreamingHttpResponse, HttpResponse
from django.core.files import File

path_file = Get_path_file()

def Generation_generation_right_not_to_withhold_pit(validated_data):
    company = CompanyAPI()
    organization = company["organizationalForm"] + ' "' + company["name"] + '"'
    list_organization = list(organization.upper())
    phone = company["contactInfo"]["phone"].upper()
    list_phone = list(phone)
    inn_company = company['inn']
    kpp_company = company['kpp']
    ogrn = company['ogrn']
    list_ogrn = list(ogrn)
    paymentAccount = company['bank']['paymentAccount'].upper()
    correspondentAccount = company['bank']['correspondentAccount'].upper()
    first_name_CEO = company['director']['fio']['firstName']
    surname_CEO = company['director']['fio']['secondName']
    patronymic_CEO = company['director']['fio']['patronymic']
    if patronymic_CEO != None and patronymic_CEO != '' and patronymic_CEO != 'string':
        CEO = surname_CEO + ' ' + first_name_CEO + ' ' + patronymic_CEO
    else:
        CEO = surname_CEO + ' ' + first_name_CEO

    INN_CEO = company['director']['inn']

    individual = IndividualAPI()

    code = validated_data['code']
    list_code = list(code)

    path_file_doc = 'document_generation_app/document_templates/right_not_to_withhold_pit.xlsx'
    doc = load_workbook(path_file_doc)
    sheet = doc.active

    list_columns_inn = ['X', 'AA', 'AD', 'AG', 'AJ',
                    'AM', 'AP', 'AS', 'AV', 'AY', 'BB', 'BH']

    row = 2
    index = 0
    for symbol in inn_company:
        for col in range(index, len(list_columns_inn)):
            cell = list_columns_inn[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            break

    list_columns_kpp = ['X', 'AA', 'AD', 'AG', 'AJ',
                    'AM', 'AP', 'AS', 'AV']
    row = 4
    index = 0
    for symbol in kpp_company:
        for col in range(index, len(list_columns_kpp)):
            cell = list_columns_kpp[index] + str(row)
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

    current_datetime = datetime.now()
    current_year = str(current_datetime.year)
    sheet['AM20'], sheet['AP20'], sheet['AS20'], sheet['AV20'] = current_year[0], current_year[1], current_year[2], \
        current_year[3]

    list_columns = ['B', 'C', 'E', 'F', 'G', 'I', 'L', 'N', 'P', 'S', 'U', 'W', 'Z', 'AC', 'AF', 'AI', 'AL',
    'AO', 'AR', 'AU']
    row = 29
    index = 0
    for symbol in CEO:
        for col in range(index, len(list_columns)):
            cell = list_columns[index] + str(row)
            if list_columns[index] == 'AU':
                sheet[f'{cell}'] = symbol
                row += 3
                index = 0
                break
            sheet[f'{cell}'] = symbol
            index += 1
            break

    list_columns_for_INN_CEO = ['F', 'G', 'I', 'L', 'N', 'P', 'S', 'U', 'W', 'Z', 'AC', 'AF']
    row = 40
    index = 0
    stop = False
    for symbol in INN_CEO:
        for col in range(index, len(list_columns_for_INN_CEO)):
            cell = list_columns_for_INN_CEO[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'AP':
                stop = True
            break
        if stop == True:
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

    row = 58
    index = 0
    for symbol in inn_company:
        for col in range(index, len(list_columns_inn)):
            cell = list_columns_inn[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            break

    list_columns_kpp = ['X', 'AA', 'AD', 'AG', 'AJ',
                    'AM', 'AP', 'AS', 'AV']
    row = 60
    index = 0
    for symbol in kpp_company:
        for col in range(index, len(list_columns_kpp)):
            cell = list_columns_kpp[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            break

    surname_worker = individual['fio']['secondName']
    first_name_worker = individual['fio']['firstName']
    patronymic_worker = individual['fio']['patronymic']
    series_passport = individual['passport']['serias']
    number_passport = individual['passport']['number']
    publisher_passport = individual['passport']['publisher']
    inn_worker = individual['inn']


    list_columns_for_surname = ['H', 'K', 'M', 'O', 'R', 'T', 'W', 'Z', 'AC', 'AF', 'AI', 'AL',
                    'AO', 'AR', 'AU', 'AX', 'BA', 'BD', 'BH']
    row = 65
    index = 0
    stop = False
    for symbol in surname_worker:
        for col in range(index, len(list_columns_for_surname)):
            cell = list_columns_for_surname[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'BH':
                stop = True
            break
        if stop == True:
            break

    list_columns_for_first_name = ['H', 'K', 'M', 'O', 'R', 'T', 'V', 'Z', 'AC', 'AF', 'AI', 'AL',
                    'AO', 'AR', 'AU', 'AX', 'BA', 'BD', 'BH']
    row = 67
    index = 0
    stop = False
    for symbol in first_name_worker:
        for col in range(index, len(list_columns_for_first_name)):
            cell = list_columns_for_first_name[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'BH':
                stop = True
            break
        if stop == True:
            break

    list_columns_for_patronymic = ['H', 'K', 'M', 'O', 'R', 'T', 'V', 'Y', 'AB', 'AE', 'AH', 'AK',
                    'AN', 'AR', 'AU', 'AX', 'BA', 'BD', 'BH']
    if patronymic_worker != None and patronymic_worker != '' and patronymic_worker != 'string':
        row = 69
        index = 0
        stop = False
        for symbol in patronymic_worker:
            for col in range(index, len(list_columns_for_patronymic)):
                cell = list_columns_for_patronymic[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'BH':
                    stop = True
                break
            if stop == True:
                break

    if series_passport != None and series_passport != '' and series_passport != 'string':
        series_and_number_passport = series_passport + ' ' + number_passport
    else:
        series_and_number_passport = number_passport
    list_columns_for_passport = ['H', 'K', 'M', 'O', 'R', 'T', 'V', 'Y', 'AB', 'AE', 'AI', 'AL',
                    'AO', 'AR', 'AU', 'AX', 'BA', 'BD', 'BH']
    row = 77
    index = 0
    stop = False
    for symbol in series_and_number_passport:
        for col in range(index, len(list_columns_for_passport)):
            cell = list_columns_for_passport[index] + str(row)
            print(cell)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'BH':
                stop = True
            break
        if stop == True:
            break

    list_columns_for_passport = ['H', 'K', 'M', 'O', 'R', 'T', 'V', 'Y', 'AB', 'AF', 'AI', 'AL',
                    'AO', 'AR', 'AU', 'AX', 'BA', 'BD', 'BH']
    row = 79
    index = 0
    stop = False
    for symbol in publisher_passport:
        for col in range(index, len(list_columns_for_passport)):
            cell = list_columns_for_passport[index] + str(row)
            print(cell)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'BH':
                stop = True
            break
        if stop == True:
            break

    global path_file
    file = ''
    if os.path.exists(path_file + '/' + 'right_not_to_withhold_pit.xlsx') == False:
        doc.save(path_file + '/' + 'right_not_to_withhold_pit.xlsx')
        file = path_file + '/' + 'right_not_to_withhold_pit.xlsx'
    else:
        i = 1
        while True:
            if os.path.exists(path_file + '/' + f'right_not_to_withhold_pit{i}.xlsx') == False:
                path = path_file + '/' + f'right_not_to_withhold_pit{i}.xlsx'
                doc.save(path)
                file = path_file + '/' + f'right_not_to_withhold_pit{i}.xlsx'
                break
            i += 1

    filename = os.path.basename(file)
    response = HttpResponse(File(open(file, 'rb')), content_type='application/vnd.ms-excel')
    response['Content-Length'] = os.path.getsize(file)
    response['Content-Disposition'] = "attachment; filename=%s" % filename
    return response
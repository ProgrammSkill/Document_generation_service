import os
from datetime import datetime
from openpyxl import *
from document_generation_app.document_generation_functions.api import CompanyAPI, IndividualAPI
from document_generation_app.document_generation_functions.functions import Date_conversion, Get_path_file

path_file = Get_path_file()

def Generation_termination_notice(validated_data):
    company = CompanyAPI()
    organization = company["organizationalForm"] + ' "' + company["name"] + '"'
    list_organization = list(organization.upper())
    phone = company["contactInfo"]["phone"].upper()
    list_phone = list(phone)
    inn = company['inn']
    kpp = company['kpp']
    list_inn_kpp = list(f'{inn}' + '/' + f'{kpp}')
    ogrn = company['ogrn']
    list_ogrn = list('ОГРН ' + ogrn)
    paymentAccount = company['bank']['paymentAccount'].upper()
    correspondentAccount = company['bank']['correspondentAccount'].upper()
    legalAddress = company['legalAddress']["name"].upper()
    list_legalAddress = list(legalAddress)
    ActualAddreses = company['ActualAddreses'][0]["name"].upper()
    CEO = 'Сталюкова Екатерина Александровна'

    individual = 'Гайназаров Кайратбек'.upper()
    # passport_series = 'AC'.upper()
    # passport_number = '4348554'.upper()

    name = validated_data['name'].upper()
    list_name = list(name)
    job_title = validated_data['job_title'].upper()
    list_job = list(job_title)
    base = validated_data['base']
    end_date = str(validated_data['end_date'].day) + '-' + str(validated_data['end_date'].month) + '-' + \
                 str(validated_data['end_date'].year)
    end_date = end_date
    initiator = validated_data['initiator']
    person = validated_data['person']

    path_file_doc = 'document_generation_app/document_templates/termination_notice.xlsx'
    doc = load_workbook(path_file_doc)
    sheet = doc.active

    list_columns = ['A', 'C', 'E', 'G', 'I', 'K', 'M', 'O', 'Q', 'S', 'U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI',
                    'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO',
                    'BQ']

    row = 11
    index = 0
    for symbol in list_name:
        for col in range(index, len(list_columns)):
            cell = list_columns[index] + str(row)
            if list_columns[index] == 'BQ':
                sheet[f'{cell}'] = symbol
                row += 2
                index = 0
                break
            sheet[f'{cell}'] = symbol
            index += 1
            break

    row = 41
    index = 0
    for symbol in list_organization:
        for col in range(index, len(list_columns)):
            cell = list_columns[index] + str(row)
            if list_columns[index] == 'BQ':
                sheet[f'{cell}'] = symbol
                row += 2
                index = 0
                break
            sheet[f'{cell}'] = symbol
            index += 1
            break

    row = 58
    index = 0
    for symbol in list_ogrn:
        for col in range(index, len(list_columns)):
            cell = list_columns[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            break

    row = 72
    index = 0
    for symbol in list_inn_kpp:
        for col in range(index, len(list_columns)):
            cell = list_columns[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            break

    row = 75
    index = 0
    for symbol in list_legalAddress:
        for col in range(index, len(list_columns)):
            cell = list_columns[index] + str(row)
            if list_columns[index] == 'BQ':
                sheet[f'{cell}'] = symbol
                row += 3
                index = 0
                break
            sheet[f'{cell}'] = symbol
            index += 1
            break

    row = 84
    index = 0
    list_columns_phone = ['U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI', 'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW',
                          'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO', 'BQ']

    for symbol in list_phone:
        for col in range(index, len(list_columns_phone)):
            cell = list_columns_phone[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            break

    row = 151
    index = 0
    for symbol in list_job:
        for col in range(index, len(list_columns)):
            cell = list_columns[index] + str(row)
            if list_columns[index] == 'BQ':
                sheet[f'{cell}'] = symbol
                row += 2
                index = 0
                break
            sheet[f'{cell}'] = symbol
            index += 1
            break

    if base == 'Трудовой договор':
        sheet['A161'] = 'X'
    elif base == 'Гражданско-правовой договор на выполнение работ (оказание услуг)':
        sheet['W161'] = 'X'

    end_date = Date_conversion(end_date)
    arr_date = end_date.split('.')
    # print(arr_date)
    # day
    sheet['AX167'] = arr_date[0][0]
    sheet['AZ167'] = arr_date[0][1]
    # month
    sheet['BC167'] = arr_date[1][0]
    sheet['BE167'] = arr_date[1][1]
    # year
    sheet['BH167'] = arr_date[2][0]
    sheet['BJ167'] = arr_date[2][1]
    sheet['BL167'] = arr_date[2][2]
    sheet['BN167'] = arr_date[2][3]

    if initiator == 'yes':
        sheet['A174'] = 'X'
    else:
        sheet['W174'] = 'X'

    sheet['AK183'] = CEO

    if person == 'person_proxy':
        try:
            full_name = validated_data['full_name']
            passportSeries = validated_data['series']
            passportNumber = validated_data['number']
            date_issue = str(validated_data['date_issue'].day) + '-' + str(validated_data['date_issue'].month) + '-' + \
                         str(validated_data['date_issue'].year)
            date_issue = date_issue
            issued_by = validated_data['issued_by']

            sheet["AE193"] = full_name
            if len(passportSeries) == 4:
                passportSeries = passportSeries[0] + passportSeries[1] + ' ' + passportSeries[2] + passportSeries[3]
            sheet["G195"] = passportSeries
            sheet["X195"] = passportNumber
            sheet["AR195"] = Date_conversion(date_issue)
            sheet["J197"] = issued_by
        except:
            pass

    else:
        sheet["AE193"] = CEO

    global path_file
    path = path_file
    if os.path.exists(path + '/' + 'termination_notice.xlsx') == False:
        doc.save(path + '/' + 'termination_notice.xlsx')
    else:
        i = 1
        while True:
            if os.path.exists(path_file + '/' + f'termination_notice{i}.xlsx') == False:
                path = path_file + '/' + f'termination_notice{i}.xlsx'
                doc.save(path)
                break
            i += 1
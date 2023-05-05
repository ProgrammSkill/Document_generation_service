import os
from datetime import datetime
from openpyxl import *
import re
from document_generation_app.document_generation_functions.api import CompanyAPI, IndividualAPI
from document_generation_app.document_generation_functions.functions import Date_conversion_from_obj_date, \
    Date_conversion, Get_path_file

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
    legalAddress = company['legalAddress']["city"] + ' г, ' + company['legalAddress']["street"] + ', ' + \
                   company['legalAddress']["house"]
    list_legalAddress = list(legalAddress.upper())
    CEO = 'Сталюкова Екатерина Александровна'

    individual = IndividualAPI()
    surname = individual['fio']['secondName'].upper()
    first_name = individual['fio']['firstName'].upper()
    patronymic = str(individual['fio']['patronymic']).upper()
    birthday = individual['birthday']
    citizenship = individual['citizenship'].upper()
    # passport
    passportSeries = individual['passport']['serias'].upper()
    passportNumber = individual['passport']['number'].upper()
    date_issue_passport = individual['passport']['dateIssue']
    publisher = individual['passport']['publisher'].upper()
    # patent
    patentSeries = individual['patent']['serias'].upper()
    patentNumber = individual['patent']['number'].upper()
    patentDateIssue = individual['patent']['dateIssue'].upper()
    patentEndDate = individual['patent']['endDate'].upper()
    patentPublisher = individual['patent']['publisher'].upper()

    name = validated_data['name'].upper()
    list_name = list(name)
    job_title = validated_data['job_title'].upper()
    list_job = list(job_title)
    base = validated_data['base']
    end_date = validated_data['end_date']
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

    list_columns_for_full_name = ['O', 'Q', 'S', 'U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI',
                    'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO',
                    'BQ']
    list_first_name = list(first_name.upper())
    row = 90
    index = 0
    stop = False
    for symbol in list_first_name:
        for col in range(index, len(list_columns_for_full_name)):
            cell = list_columns_for_full_name[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'BQ':
                stop = True
            break
        if stop == True:
            break

    list_surname = list(surname.upper())
    row = 92
    index = 0
    stop = False
    for symbol in list_surname:
        for col in range(index, len(list_columns_for_full_name)):
            cell = list_columns_for_full_name[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'BQ':
                stop = True
            break
        if stop == True:
            break

    if patronymic != None and patronymic != '':
        list_patronymic = list(patronymic.upper())
        row = 94
        index = 0
        stop = False
        for symbol in list_patronymic:
            for col in range(index, len(list_columns_for_full_name)):
                cell = list_columns_for_full_name[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                if col == 'BQ':
                    stop = True
                break
            if stop == True:
                break

    list_columns_for_citizenship = ['Q', 'S', 'U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI',
                    'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO',
                    'BQ']
    row = 97
    index = 0
    stop = False
    for symbol in citizenship:
        for col in range(index, len(list_columns_for_citizenship)):
            cell = list_columns_for_citizenship[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'BQ':
                stop = True
            break
        if stop == True:
            break

    birthday = re.search("([0-9]{4}\-[0-9]{2}\-[0-9]{2})", birthday).group(1)
    list_birthday = birthday.split('-')
    #day
    sheet['R104'],  sheet['T104'] = list_birthday[2][0], list_birthday[2][1]
    #month
    sheet['W104'], sheet['Y104'] = list_birthday[1][0], list_birthday[1][1]
    #year
    sheet['AB104'], sheet['AD104'] = list_birthday[0][0], list_birthday[0][1]
    sheet['AF104'], sheet['AH104'] = list_birthday[0][2], list_birthday[0][3]

    # ============= passport record =============
    list_passportSeries = list(passportSeries.upper())
    list_columns_for_passportSeries = ['G', 'I', 'K', 'M', 'O', 'Q', 'S']
    row = 110
    index = 0
    stop = False
    for symbol in list_passportSeries:
        for col in range(index, len(list_columns_for_passportSeries)):
            cell = list_columns_for_passportSeries[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'S':
                stop = True
            break
        if stop == True:
            break

    list_passportNumber = passportNumber.upper()
    list_columns_for_passportNumber = ['Z', 'AB', 'AD', 'AF', 'AH', 'AJ', 'AL', 'AN', 'AP']
    row = 110
    index = 0
    stop = False
    for symbol in list_passportNumber:
        for col in range(index, len(list_columns_for_passportNumber)):
            cell = list_columns_for_passportNumber[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'AP':
                stop = True
            break
        if stop == True:
            break

    date_issue_passport = re.search("([0-9]{4}\-[0-9]{2}\-[0-9]{2})", date_issue_passport).group(1)
    list_date_issue_passport = date_issue_passport.split('-')
    #day
    sheet['BA110'],  sheet['BC110'] = list_date_issue_passport[2][0], list_date_issue_passport[2][1]
    #month
    sheet['BF110'], sheet['BH110'] = list_date_issue_passport[1][0], list_date_issue_passport[1][1]
    #year
    sheet['BK110'], sheet['BM110'] = list_date_issue_passport[0][0], list_date_issue_passport[0][1]
    sheet['BO110'], sheet['BQ110'] = list_date_issue_passport[0][2], list_date_issue_passport[0][3]


    list_columns_for_publisher = ['O', 'Q', 'S', 'U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI',
     'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO',
     'BQ']
    list_columns_for_publisher_last = ['AD', 'AF', 'AH', 'AJ', 'AL', 'AN', 'AP', 'AR', 'AT', 'AV', 'AX', 'AZ',
     'BB']
    row = 113
    index_publisher_last = 0
    stop = False
    for symbol in publisher:
        if row == 117:
            for col in range(index_publisher_last, len(list_columns_for_publisher_last)):
                cell = list_columns_for_publisher_last[index_publisher_last] + str(row)
                sheet[f'{cell}'] = symbol
                index_publisher_last += 1
                if col == 'BB':
                    stop = True
                break
            if stop == True:
                break
        else:
            for col in range(index, len(list_columns_for_publisher)):
                cell = list_columns_for_publisher[index] + str(row)
                if list_columns_for_publisher[index] == 'BQ':
                    sheet[f'{cell}'] = symbol
                    row += 2
                    index = 0
                    break
                sheet[f'{cell}'] = symbol
                index += 1
                break
    # =======================================================================

    # ============= patent record =============
    list_columns_for_patentSeries = ['G', 'I', 'K', 'M', 'O', 'Q', 'S']
    row = 125
    index = 0
    stop = False
    for symbol in patentSeries:
        for col in range(index, len(list_columns_for_patentSeries)):
            cell = list_columns_for_patentSeries[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'S':
                stop = True
            break
        if stop == True:
            break

    list_columns_for_patentNumber = ['Z', 'AB', 'AD', 'AF', 'AH', 'AJ', 'AL', 'AN', 'AP']
    row = 125
    index = 0
    stop = False
    for symbol in patentNumber:
        for col in range(index, len(list_columns_for_patentNumber)):
            cell = list_columns_for_patentNumber[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'AP':
                stop = True
            break
        if stop == True:
            break

    patentDateIssue = re.search("([0-9]{4}\-[0-9]{2}\-[0-9]{2})", patentDateIssue).group(1)
    list_date_issue_patent = patentDateIssue.split('-')
    # day
    sheet['BA125'], sheet['BC125'] = list_date_issue_patent[2][0], list_date_issue_patent[2][1]
    # month
    sheet['BF125'], sheet['BH125'] = list_date_issue_patent[1][0], list_date_issue_patent[1][1]
    # year
    sheet['BK125'], sheet['BM125'] = list_date_issue_patent[0][0], list_date_issue_patent[0][1]
    sheet['BO125'], sheet['BQ125'] = list_date_issue_patent[0][2], list_date_issue_patent[0][3]

    list_columns_for_publisher = ['Q', 'S', 'U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI',
                    'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO',
                    'BQ']
    list_columns_for_publisher_last = ['A', 'C', 'E', 'G', 'I', 'K', 'M', 'O', 'Q', 'S', 'U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI',
                    'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO',
                    'BQ']
    row = 133
    index_publisher_last = 0
    stop = False
    for symbol in publisher:
        if row == 129:
            for col in range(index_publisher_last, len(list_columns_for_publisher_last)):
                cell = list_columns_for_publisher_last[index_publisher_last] + str(row)
                sheet[f'{cell}'] = symbol
                index_publisher_last += 1
                if col == 'BQ':
                    stop = True
                break
            if stop == True:
                break
        else:
            for col in range(index, len(list_columns_for_publisher)):
                cell = list_columns_for_publisher[index] + str(row)
                if list_columns_for_publisher[index] == 'BQ':
                    sheet[f'{cell}'] = symbol
                    row += 2
                    index = 0
                    break
                sheet[f'{cell}'] = symbol
                index += 1
                break

    # day
    sheet['P131'], sheet['R131'] = list_date_issue_patent[2][0], list_date_issue_patent[2][1]
    # month
    sheet['U131'], sheet['W131'] = list_date_issue_patent[1][0], list_date_issue_patent[1][1]
    # year
    sheet['Z131'], sheet['AB131'] = list_date_issue_patent[0][0], list_date_issue_patent[0][1]
    sheet['AD131'], sheet['AF131'] = list_date_issue_patent[0][2], list_date_issue_patent[0][3]

    patentEndDate = re.search("([0-9]{4}\-[0-9]{2}\-[0-9]{2})", patentEndDate).group(1)
    list_end_date_patent = patentEndDate.split('-')
    # day
    sheet['AL131'], sheet['AN131'] = list_end_date_patent[2][0], list_end_date_patent[2][1]
    # month
    sheet['AQ131'], sheet['AS131'] = list_end_date_patent[1][0], list_end_date_patent[1][1]
    # year
    sheet['AV131'], sheet['AX131'] = list_end_date_patent[0][0], list_end_date_patent[0][1]
    sheet['AZ131'], sheet['BB131'] = list_end_date_patent[0][2], list_end_date_patent[0][3]
    # ===========================================

    citizenship = individual['citizenship']
    if citizenship == 'Киргизия' or citizenship == 'Армения' or citizenship == 'Казахстан' or citizenship == 'Беларусь':
        name_international_agreement = "П. 1, СТАТЬИ 97, ДОГОВОРА О ЕВРАЗИЙСКОМ ЭКОНОМИЧЕСКОМ СОЮЗЕ ОТ 29.05.2014 (В РЕД. ОТ 08.05.2015)"
        list_international_agreement = list(name_international_agreement)
        row = 140
        index = 0
        for symbol in list_international_agreement:
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

    end_date = Date_conversion_from_obj_date(end_date)
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
        sheet['E174'] = 'X'
    else:
        sheet['M174'] = 'X'

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
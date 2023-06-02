import os
from openpyxl import *
import re
from document_generation_app.document_generation_functions.api import CompanyAPI, IndividualAPI
from document_generation_app.document_generation_functions.functions import Date_conversion_from_obj_date, \
    Date_conversion, Get_path_file
from django.http import FileResponse, HttpResponse, StreamingHttpResponse, HttpResponse
from django.core.files import File

path_file = Get_path_file()

def Generation_notice_conclusion(validated_data):
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
    okved = company['okved']
    legalAddress = company['legalAddress']["city"] + ' г, ' + company['legalAddress']["street"] + ', ' + \
                   company['legalAddress']["house"]
    list_legalAddress = list(legalAddress.upper())

    first_name_CEO = company['director']['fio']['firstName']
    surname_CEO = company['director']['fio']['secondName']
    patronymic_CEO = company['director']['fio']['patronymic']
    if patronymic_CEO != None and patronymic_CEO != '' and patronymic_CEO != 'string':
        CEO = surname_CEO + ' ' + first_name_CEO + ' ' + patronymic_CEO
    else:
        CEO = surname_CEO + ' ' + first_name_CEO

    individual = IndividualAPI()
    surname = individual['fio']['secondName']
    first_name = individual['fio']['firstName']
    patronymic = individual['fio']['patronymic']
    citizenship = individual['citizenship']
    birthplace = individual['birthplace']['city'].upper()
    # passport
    passportSeries = individual['passport']['serias'].upper()
    passportNumber = individual['passport']['number'].upper()
    date_issue_passport = individual['passport']['dateIssue']
    publisher = individual['passport']['publisher'].upper()

    name_territorial_body = validated_data['name_territorial_body'].upper()
    job_title = validated_data['job_title'].upper()
    list_job = list(job_title)
    base = validated_data['base']
    start_date = validated_data['start_date']
    address = validated_data['address'].upper()
    list_address = list(address)
    person = validated_data['person']

    path_file_doc = 'document_generation_app/document_templates/notice_conclusion.xlsx'
    doc = load_workbook(path_file_doc)
    sheet = doc.active

    list_columns = ['A', 'C', 'E', 'G', 'I', 'K', 'M', 'O', 'Q', 'S', 'U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI', 'AK',
                    'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO', 'BQ']

    row = 11
    index = 0
    for symbol in name_territorial_body:
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

    list_columns_for_okved = ['AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO', 'BQ']
    row = 36
    index = 0
    stop = False
    for symbol in okved:
        for col in range(index, len(list_columns_for_okved)):
            cell = list_columns_for_okved[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'BQ':
                stop = True
            break
        if stop == True:
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

    row = 73
    index = 0
    for symbol in list_inn_kpp:
        for col in range(index, len(list_columns)):
            cell = list_columns[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            break

    row = 76
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

    row = 85
    index = 0
    list_columns_phone = ['U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI', 'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY',
                          'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO', 'BQ']
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
    row = 91
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
    row = 93
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

    if patronymic != None and patronymic != '' and patronymic != 'string':
        list_patronymic = list(patronymic.upper())
        row = 95
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

    list_citizenship = citizenship.upper()
    list_columns_for_citizenship = ['Q', 'S', 'U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI',
                    'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO',
                    'BQ']
    row = 98
    index = 0
    stop = False
    for symbol in list_citizenship:
        for col in range(index, len(list_columns_for_citizenship)):
            cell = list_columns_for_citizenship[index] + str(row)
            sheet[f'{cell}'] = symbol
            index += 1
            if col == 'BQ':
                stop = True
            break
        if stop == True:
            break

    list_columns_for_citizenship = ['W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI',
                    'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO',
                    'BQ']
    row = 100
    index = 0
    index_citizenship_last = 0
    stop = False
    for symbol in birthplace:
        if row == 103:
            for col in range(index_citizenship_last, len(list_columns)):
                cell = list_columns[index_citizenship_last] + str(row)
                sheet[f'{cell}'] = symbol
                index_citizenship_last += 1
                if col == 'BQ':
                    stop = True
                break
            if stop == True:
                break
        else:
            for col in range(index, len(list_columns_for_citizenship)):
                cell = list_columns_for_citizenship[index] + str(row)
                if list_columns_for_citizenship[index] == 'BQ':
                    sheet[f'{cell}'] = symbol
                    row += 3
                    index = 0
                    break
                sheet[f'{cell}'] = symbol
                index += 1
                break

    birthday = individual['birthday']
    birthday = re.search("([0-9]{4}\-[0-9]{2}\-[0-9]{2})", birthday).group(1)
    list_birthday = birthday.split('-')
    #day
    sheet['R105'],  sheet['T105'] = list_birthday[2][0], list_birthday[2][1]
    #month
    sheet['W105'], sheet['Y105'] = list_birthday[1][0], list_birthday[1][1]
    #year
    sheet['AB105'], sheet['AD105'] = list_birthday[0][0], list_birthday[0][1]
    sheet['AF105'], sheet['AH105'] = list_birthday[0][2], list_birthday[0][3]

    list_passportSeries = list(passportSeries.upper())
    list_columns_for_passportSeries = ['G', 'I', 'K', 'M', 'O', 'Q', 'S']
    row = 111
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
    row = 111
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
    sheet['BA111'],  sheet['BC111'] = list_date_issue_passport[2][0], list_date_issue_passport[2][1]
    #month
    sheet['BF111'], sheet['BH111'] = list_date_issue_passport[1][0], list_date_issue_passport[1][1]
    #year
    sheet['BK111'], sheet['BM111'] = list_date_issue_passport[0][0], list_date_issue_passport[0][1]
    sheet['BO111'], sheet['BQ111'] = list_date_issue_passport[0][2], list_date_issue_passport[0][3]


    list_columns_for_publisher = ['O', 'Q', 'S', 'U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI',
     'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO', 'BQ']
    list_columns_for_publisher_last = ['AD', 'AF', 'AH', 'AJ', 'AL', 'AN', 'AP', 'AR', 'AT', 'AV', 'AX', 'AZ', 'BB']
    row = 114
    index = 0
    index_publisher_last = 0
    stop = False
    for symbol in publisher:
        if row == 118:
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
    if individual['patent'] != None:
        # patent
        patentSeries = individual['patent']['serias'].upper()
        patentNumber = individual['patent']['number'].upper()
        patentDateIssue = individual['patent']['dateIssue'].upper()
        patentEndDate = individual['patent']['endDate'].upper()
        patentPublisher = individual['patent']['publisher'].upper()
        list_columns_for_patentSeries = ['G', 'I', 'K', 'M', 'O', 'Q', 'S']
        row = 131
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

        list_columns_for_patentNumber = ['X', 'Z', 'AB', 'AD', 'AF', 'AH', 'AJ', 'AL', 'AN', 'AP']
        row = 131
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
        sheet['BA131'], sheet['BC131'] = list_date_issue_patent[2][0], list_date_issue_patent[2][1]
        # month
        sheet['BF131'], sheet['BH131'] = list_date_issue_patent[1][0], list_date_issue_patent[1][1]
        # year
        sheet['BK131'], sheet['BM131'] = list_date_issue_patent[0][0], list_date_issue_patent[0][1]
        sheet['BO131'], sheet['BQ131'] = list_date_issue_patent[0][2], list_date_issue_patent[0][3]

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
            if row == 135:
                for col in range(index_publisher_last, len(list_columns_for_publisher_last)):
                    cell = list_columns_for_publisher_last[index_publisher_last] + str(row)
                    print(cell)
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
        sheet['P137'], sheet['R137'] = list_date_issue_patent[2][0], list_date_issue_patent[2][1]
        # month
        sheet['U137'], sheet['W137'] = list_date_issue_patent[1][0], list_date_issue_patent[1][1]
        # year
        sheet['Z137'], sheet['AB137'] = list_date_issue_patent[0][0], list_date_issue_patent[0][1]
        sheet['AD137'], sheet['AF137'] = list_date_issue_patent[0][2], list_date_issue_patent[0][3]

        patentEndDate = re.search("([0-9]{4}\-[0-9]{2}\-[0-9]{2})", patentEndDate).group(1)
        list_end_date_patent = patentEndDate.split('-')
        # day
        sheet['AL137'], sheet['AN137'] = list_end_date_patent[2][0], list_end_date_patent[2][1]
        # month
        sheet['AQ137'], sheet['AS137'] = list_end_date_patent[1][0], list_end_date_patent[1][1]
        # year
        sheet['AV137'], sheet['AX137'] = list_end_date_patent[0][0], list_end_date_patent[0][1]
        sheet['AZ137'], sheet['BB137'] = list_end_date_patent[0][2], list_end_date_patent[0][3]
    # ===========================================
    else:
        citizenship = individual['citizenship']
        if citizenship == 'Киргизия' or citizenship == 'Армения' or citizenship == 'Казахстан' or citizenship == 'Беларусь':
            name_international_agreement = "П. 1, СТАТЬИ 97, ДОГОВОРА О ЕВРАЗИЙСКОМ ЭКОНОМИЧЕСКОМ СОЮЗЕ ОТ 29.05.2014 (В РЕД. ОТ 08.05.2015)"
            row = 146
            index = 0
            for symbol in name_international_agreement:
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

    row = 157
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

    if 'Трудовой договор' in base:
        sheet['A167'] = 'X'
    elif 'Гражданско-правовой договор на выполнение работ (оказание услуг)' in base:
        sheet['W167'] = 'X'

    start_date = Date_conversion_from_obj_date(start_date)
    start_date = Date_conversion(start_date)
    arr_date = start_date.split('.')
    # day
    sheet['AX172'] = arr_date[0][0]
    sheet['AZ172'] = arr_date[0][1]
    # month
    sheet['BC172'] = arr_date[1][0]
    sheet['BE172'] = arr_date[1][1]
    # year
    sheet['BH172'] = arr_date[2][0]
    sheet['BJ172'] = arr_date[2][1]
    sheet['BL172'] = arr_date[2][2]
    sheet['BN172'] = arr_date[2][3]

    row = 178
    index = 0
    for symbol in list_address:
        for col in range(index, len(list_columns)):
            cell = list_columns[index] + str(row)
            if list_columns[index] == 'BQ':
                sheet[f'{cell}'] = symbol
                if row == 180:
                    row += 3
                else:
                    row += 2
                index = 0
                break
            sheet[f'{cell}'] = symbol
            index += 1
            break

    sheet['AK191'] = CEO
    print(person)
    if person == 'Человек, который подаёт документы по доверенности':
        try:
            full_name = validated_data['full_name']
            passportSeries = validated_data['series']
            passportNumber = validated_data['number']
            date_issue = validated_data['date_issue']
            issued_by = validated_data['issued_by']

            sheet["AE201"] = full_name
            if len(passportSeries) == 4:
                passportSeries = passportSeries[0] + passportSeries[1] + ' ' + passportSeries[2] + passportSeries[3]
            sheet["G203"] = passportSeries
            sheet["X203"] = passportNumber
            date_issue = Date_conversion_from_obj_date(date_issue)
            sheet["AR203"] = Date_conversion(date_issue)
            sheet["J205"] = issued_by
        except:
            pass

    else:
        sheet["AE201"] = CEO
        passportSeries = company["director"]["passport"]['serias']
        passportNumber = company["director"]["passport"]['number']
        date_issue = company["director"]["passport"]['date_issue']
        issued_by = company["director"]["passport"]['issued_by']

        if len(passportSeries) == 4:
            passportSeries = passportSeries[0] + passportSeries[1] + ' ' + passportSeries[2] + passportSeries[3]
        sheet["G203"] = passportSeries
        sheet["X203"] = passportNumber
        date_issue = Date_conversion_from_obj_date(date_issue)
        sheet["AR203"] = Date_conversion(date_issue)
        sheet["J205"] = issued_by

    global path_file
    file = ''
    if os.path.exists(path_file + '/' + 'notice_conclusion.xlsx') == False:
        doc.save(path_file + '/' + 'notice_conclusion.xlsx')
        file = path_file + '/' + 'notice_conclusion.xlsx'
    else:
        i = 1
        while True:
            if os.path.exists(path_file + '/' + f'notice_conclusiont{i}.xlsx') == False:
                path = path_file + '/' + f'notice_conclusion{i}.xlsx'
                doc.save(path)
                file = path_file + '/' + f'notice_conclusion{i}.xlsx'
                break
            i += 1

    filename = os.path.basename(file)
    response = HttpResponse(File(open(file, 'rb')), content_type='application/vnd.ms-excel')
    response['Content-Length'] = os.path.getsize(file)
    response['Content-Disposition'] = "attachment; filename=%s" % filename
    return response

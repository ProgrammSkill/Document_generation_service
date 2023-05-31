import os
from docxtpl import DocxTemplate
from number_to_string import get_string_by_number
from document_generation_app.document_generation_functions.api import CompanyAPI, IndividualAPI
from document_generation_app.document_generation_functions.functions import Date_conversion_from_obj_date, \
    Date_conversion, Get_path_file, SurnameDeclension, FirstNameDeclension, LastNameDeclension, CountryDeclination
import re

path_file = Get_path_file()


def Generation_employment_contract_document(validated_data):
    company = CompanyAPI()
    organization = company["organizationalForm"] + ' "' + company["name"] + '"'
    phone = company["contactInfo"]["phone"]
    inn = company['inn']
    kpp = company['kpp']
    paymentAccount = company['bank']['paymentAccount']
    correspondentAccount = company['bank']['correspondentAccount']
    city = company['legalAddress']["city"]
    legalAddress = company['legalAddress']['postalCode'] + ', ' + city + ' г, ' + company['legalAddress']["street"] + \
                   ', ' + company['legalAddress']["house"]
    ActualAddreses = company['ActualAddreses'][0]['postalCode'] + ', ' + company['ActualAddreses'][0]["city"] + \
        ' г, ' + company['ActualAddreses'][0]["street"] + ', ' + company['ActualAddreses'][0]["house"]
    BIC = company['bank']['bankId']
    nameBank = company['bank']['nameBank']


    first_name_CEO = company['director']['fio']['firstName']
    surname_CEO = company['director']['fio']['secondName']
    patronymic_CEO = company['director']['fio']['patronymic']

    declension_first_name_CEO = FirstNameDeclension(first_name_CEO)
    declension_surname_CEO = SurnameDeclension(surname_CEO)
    declension_patronymic_CEO = LastNameDeclension(patronymic_CEO)

    if patronymic_CEO != None and patronymic_CEO != '' and patronymic_CEO != 'string':
        CEO = surname_CEO + ' ' + first_name_CEO + ' ' + patronymic_CEO
        CEO_declension = declension_surname_CEO + ' ' + declension_first_name_CEO + ' ' + declension_patronymic_CEO
        surname_initials_CEO = declension_surname_CEO + ' ' + first_name_CEO[0] + '.' + patronymic_CEO[0] + '.'
    else:
        CEO = surname_CEO + ' ' + first_name_CEO
        CEO_declension = declension_surname_CEO + ' ' + declension_first_name_CEO
        surname_initials_CEO = declension_surname_CEO + ' ' + first_name_CEO[0] + '.'

    individual = IndividualAPI()
    surname = individual['fio']['secondName']
    name = individual['fio']['firstName']
    patronymic = individual['fio']['patronymic']
    birthDay = individual['birthday']
    citizenship = CountryDeclination(individual['citizenship']).upper()
    birthDay = re.search("([0-9]{4}\-[0-9]{2}\-[0-9]{2})", birthDay).group(1).split('-')
    birthDay = birthDay[2] + '.' + birthDay[1] + '.' + birthDay[0]

    passportSeries = individual['passport']['serias']
    passportNumber = individual['passport']['number']

    if passportSeries != '' and passportSeries != None or passportSeries != 'string':
        passport = passportSeries + passportNumber
    else:
        passport = passportNumber

    patentSeries = individual['patent']['serias']
    patentNumber = individual['patent']['number']
    patent = patentSeries + patentNumber


    dateIssuePassport = individual['passport']['dateIssue']
    dateIssuePassport = re.search("([0-9]{4}\-[0-9]{2}\-[0-9]{2})", dateIssuePassport).group(1).split('-')
    dateIssuePassport = dateIssuePassport[2] + '.' + dateIssuePassport[1] + '.' + dateIssuePassport[0]
    endDatePassport = individual['passport']['endDate']
    endDatePassport = re.search("([0-9]{4}\-[0-9]{2}\-[0-9]{2})", endDatePassport).group(1).split('-')
    endDatePassport = endDatePassport[2] + '.' + endDatePassport[1] + '.' + endDatePassport[0]

    dateIssuePatent = individual['patent']['dateIssue']
    dateIssuePatent = re.search("([0-9]{4}\-[0-9]{2}\-[0-9]{2})", dateIssuePatent).group(1).split('-')
    dateIssuePatent = dateIssuePatent[2] + '.' + dateIssuePatent[1] + '.' + dateIssuePatent[0]

    if patronymic != None and patronymic != '' and patronymic != 'string':
        full_name_worker = surname + ' ' + name + ' ' + patronymic
    else:
        full_name_worker = surname + ' ' + name

    number = validated_data['number']
    job_title = validated_data['job_title']
    salary = validated_data['salary']
    contract_type = validated_data['contract_type']
    start_date = Date_conversion_from_obj_date(validated_data['start_date'])

    if contract_type == 'perpetual':
        startDateWordMonth = Date_conversion(start_date, 'word_month')
        date_content = f' и является бессрочным Дата начала работы по настоящему Договору: {startDateWordMonth}'
    else:
        startDateWordMonth = Date_conversion(start_date, 'word_month')
        end_date = Date_conversion_from_obj_date(validated_data['end_date_urgent'])
        endDateWordMonth = Date_conversion(end_date, 'word_month')
        cause = validated_data['cause']
        date_content = f'. Настоящий трудовой договор является срочным, заключается на срок с {startDateWordMonth} по {endDateWordMonth} Обстоятельства (причины), послужившие основанием для заключения срочного трудового договора, - {cause}'

    start_time = str(validated_data['start_time'])
    if start_time[0] == "0":
        start_time = start_time[1:]
    end_time = str(validated_data['end_time'])
    if end_time[0] == "0":
        end_time = end_time[1:]

    start_time = start_time.split(':')
    start_time = start_time[0] + ':' + start_time[1]
    end_time = end_time.split(':')
    end_time = end_time[0] + ':' + end_time[1]

    textSalary = get_string_by_number(salary).replace(' рублей 00 копеек', '', 1)
    path_file_doc = 'document_generation_app/document_templates/employment_contract.docx'
    doc = DocxTemplate(path_file_doc)

    context = {
        'organization': organization,
        'surnameInitialsCEO': surname_initials_CEO,
        'number': number,
        'job_title': job_title,
        'salary': salary,
        'textSalary': textSalary,
        'startDateQuotes': Date_conversion(start_date, 'quotes'),
        'startDateWordMonth': Date_conversion(start_date, 'word_month'),
        'startDateStandart': Date_conversion(start_date),
        'dateContent': date_content,
        'startTime': start_time,
        'endTime': end_time,
        'inn': inn,
        'kpp': kpp,
        'phone': phone,
        'city': city,
        'legalAddress': legalAddress,
        'ActualAddreses': ActualAddreses,
        'paymentAccount': paymentAccount,
        'correspondentAccount': correspondentAccount,
        'ceoDeclension': CEO_declension,
        'CEO': CEO,
        'fullName': full_name_worker,
        'citizenship': citizenship,
        'birthDay': birthDay,
        'BIC': BIC,
        'nameBank': nameBank,
        'dateIssuePassport': dateIssuePassport,
        'endDatePassport': endDatePassport,
        'dateIssuePatent': dateIssuePatent,
        'passport': passport,
        'patent': patent
    }

    doc.render(context)

    global path_file
    path = path_file
    if os.path.exists(path + '/' + 'employment_contract.docx') == False:
        doc.save(path + '/' + 'employment_contract.docx')
    else:
        i = 1
        while True:
            if os.path.exists(path_file + '/' + f'employment_contract{i}.docx') == False:
                path = path_file + '/' + f'employment_contract{i}.docx'
                doc.save(path)
                break
            i += 1
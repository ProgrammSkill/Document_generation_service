import os
import requests
from django.shortcuts import render, redirect
from docxtpl import DocxTemplate
from datetime import datetime
import pytz
from number_to_string import get_string_by_number
from document_generation_app.document_generation_functions.api import CompanyAPI, IndividualAPI
from document_generation_app.document_generation_functions.functions import Date_conversion_from_obj_date, Date_conversion, Get_path_file

path_file = Get_path_file()


def Generation_employment_contract_document(validated_data):
    company = CompanyAPI()
    organization = company["organizationalForm"] + ' "' + company["name"] + '"'
    phone = company["contactInfo"]["phone"]
    inn = company['inn']
    kpp = company['kpp']
    ogrn = company['ogrn']
    paymentAccount = company['bank']['paymentAccount']
    correspondentAccount = company['bank']['correspondentAccount']
    legalAddress = company['legalAddress']["name"]
    ActualAddreses = company['ActualAddreses'][0]["name"]
    CEO = 'Сталюковой Екатерина Александровны'

    individual = 'Гайназаров Кайратбек'
    passport_series = 'AC'
    passport_number = '4348554'

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
        end_date = Date_conversion_from_obj_date(validated_data['end_date'])
        endDateWordMonth = Date_conversion(end_date, 'word_month')
        cause = validated_data['cause']
        date_content = f'. Настоящий трудовой договор является срочным, заключается на срок {startDateWordMonth} по {endDateWordMonth} Обстоятельства (причины), послужившие основанием для заключения срочного трудового договора, - {cause}'

    start_time = str(validated_data['start_time'])
    if start_time[0] == "0":
        start_time = start_time[1:]
    end_time = str(validated_data['end_date_urgent'])
    if end_time[0] == "0":
        end_time = end_time[1:]

    textSalary = get_string_by_number(salary).replace(' рублей 00 копеек', '', 1)
    path_file_doc = 'document_generation_app/document_templates/employment_contract.docx'
    doc = DocxTemplate(path_file_doc)

    context = {
        'organization': organization,
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
        'legalAddress': legalAddress,
        'ActualAddreses': ActualAddreses,
        'paymentAccount': paymentAccount,
        'correspondentAccount': correspondentAccount,
        'CEO': CEO,
        'individual': individual
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
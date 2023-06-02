import os
from docxtpl import DocxTemplate
from datetime import datetime
from document_generation_app.document_generation_functions.api import CompanyAPI, IndividualAPI
from document_generation_app.document_generation_functions.functions import Date_conversion_from_obj_date, \
    Date_conversion, Get_path_file
from django.http import FileResponse, HttpResponse, StreamingHttpResponse, HttpResponse
from django.core.files import File

path_file = Get_path_file()

def Generation_suspension_order(validated_data):
    company = CompanyAPI()
    organization = company["organizationalForm"] + ' "' + company["name"] + '"'
    phone = company["contactInfo"]["phone"]
    inn = company['inn']
    kpp = company['kpp']
    paymentAccount = company['bank']['paymentAccount']
    correspondentAccount = company['bank']['correspondentAccount']
    legalAddress = company['legalAddress']['postalCode'] + ', ' + company['legalAddress']['city'] + ' г, ' + company['legalAddress']["street"] + \
                   ', ' + company['legalAddress']["house"]
    ActualAddreses = company['ActualAddreses'][0]['postalCode'] + ', ' + company['ActualAddreses'][0]["city"] + \
        ' г, ' + company['ActualAddreses'][0]["street"] + ', ' + company['ActualAddreses'][0]["house"]
    first_name_CEO = company['director']['fio']['firstName']
    surname_CEO = company['director']['fio']['secondName']
    patronymic_CEO = company['director']['fio']['patronymic']

    if patronymic_CEO != None and patronymic_CEO != '' and patronymic_CEO != 'string':
        CEO = surname_CEO + ' ' + first_name_CEO + ' ' + patronymic_CEO
    else:
        CEO = surname_CEO + ' ' + first_name_CEO

    individual = IndividualAPI()
    surname = individual['fio']['secondName']
    name = individual['fio']['firstName']
    patronymic = individual['fio']['patronymic']

    if patronymic != None and patronymic != '' and patronymic != 'string':
        full_name_worker = surname + ' ' + name + ' ' + patronymic
    else:
        full_name_worker = surname + ' ' + name


    number = validated_data['number']
    start_date = Date_conversion_from_obj_date(validated_data['start_date'])
    reason_suspension = validated_data['reason_suspension']
    first_point_performer = validated_data['first_point_performer']
    second_point_performer = validated_data['second_point_performer']

    path_file_doc = 'document_generation_app/document_templates/suspension_order.docx'
    doc = DocxTemplate(path_file_doc)

    context = {
        'organization': organization,
        'number': number,
        'startDateQuotes': Date_conversion(start_date, 'quotes'),
        'startDateWordMonth': Date_conversion(start_date, 'word_month'),
        'startDateStandart': Date_conversion(start_date),
        'reasonSuspension': reason_suspension,
        'first': first_point_performer,
        'second': second_point_performer,
        'inn': inn,
        'kpp': kpp,
        'phone': phone,
        'legalAddress': legalAddress,
        'ActualAddreses': ActualAddreses,
        'paymentAccount': paymentAccount,
        'correspondentAccount': correspondentAccount,
        'CEO': CEO,
        'individual': full_name_worker
    }

    doc.render(context)

    global path_file
    file = ''
    if os.path.exists(path_file + '/' + 'suspension_order.docx') == False:
        doc.save(path_file + '/' + 'suspension_order.docx')
        file = path_file + '/' + 'suspension_order.docx'
    else:
        i = 1
        while True:
            if os.path.exists(path_file + '/' + f'suspension_order{i}.docx') == False:
                path = path_file + '/' + f'suspension_order{i}.docx'
                doc.save(path)
                file = path_file + '/' + f'suspension_order{i}.docx'
                break
            i += 1

    filename = os.path.basename(file)
    response = HttpResponse(File(open(file, 'rb')), content_type='application/vnd.ms-exce')
    response['Content-Length'] = os.path.getsize(file)
    response['Content-Disposition'] = "attachment; filename=%s" % filename
    return response
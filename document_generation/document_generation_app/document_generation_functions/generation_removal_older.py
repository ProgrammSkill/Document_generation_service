import os
from docxtpl import DocxTemplate
from datetime import datetime
from document_generation_app.document_generation_functions.api import CompanyAPI, IndividualAPI
from document_generation_app.document_generation_functions.functions import Date_conversion, Get_path_file

path_file = Get_path_file()

def Generation_removal_older(request):
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

    number = request.POST.get('number')
    obj_start_date = datetime.strptime(request.POST.get('start_date'), '%Y-%m-%d')
    start_date = obj_start_date.strftime("%d-%m-%Y")

    path_file_doc = 'document_generation_app/document_templates/removal_order.docx'
    doc = DocxTemplate(path_file_doc)

    context = {
        'organization': organization,
        'number': number,
        'startDateQuotes': Date_conversion(start_date, 'quotes'),
        'startDateWordMonth': Date_conversion(start_date, 'word_month'),
        'startDateStandart': Date_conversion(start_date),
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
    if os.path.exists(path + '/' + 'removal_order.docx') == False:
        doc.save(path + '/' + 'removal_order.docx')
    else:
        i = 1
        while True:
            if os.path.exists(path_file + '/' + f'removal_order{i}.docx') == False:
                path = path_file + '/' + f'removal_order{i}.docx'
                doc.save(path)
                break
            i += 1
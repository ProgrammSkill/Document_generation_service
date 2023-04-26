import os
from docxtpl import DocxTemplate
from datetime import datetime
from document_generation_app.document_generation_functions.api import CompanyAPI, IndividualAPI
from document_generation_app.document_generation_functions.functions import Date_conversion, Get_path_file

path_file = Get_path_file()

def Generation_GPH_contract(validated_data):
    company = CompanyAPI()
    paymentAccount_organization = company['bank']['paymentAccount']
    organization = company["organizationalForm"] + " " + company["name"]
    legalAddress = company['legalAddress']["name"]
    organization_inn = company['inn']
    organization_kpp = company['kpp']
    organization_ogrn = company['ogrn']

    CEO = 'Сталюковой Екатерина Александровны'

    individual = IndividualAPI()
    surname = individual['surname']
    name = individual['name']
    patronymic = individual['patronymic']
    passportSeries = individual['passport']['series']
    passportNumber = individual['passport']['number']
    temporary_registration = individual['temporary_registration']

    if patronymic != None and patronymic != '':
        full_name = surname + ' ' + name + ' ' + patronymic
    else:
        full_name = surname + ' ' + name

    arr_value_table = []

    number = validated_data['number']
    start_date = str(validated_data['start_date'].day) + '-' + str(validated_data['start_date'].month) + '-' + \
                 str(validated_data['start_date'].year)
    start_date = Date_conversion(start_date, 'word_month')
    end_date = str(validated_data['end_date'].day) + '-' + str(validated_data['end_date'].month) + '-' + \
                 str(validated_data['end_date'].year)
    end_date = Date_conversion(end_date, 'word_month')
    address = validated_data['address']
    # table_doc = request.POST.get('table_doc')

    path_file_doc = 'document_generation_app/document_templates/GPC_agreement.docx'
    doc = DocxTemplate(path_file_doc)

    context = {
        'number': number,
        'startDate': start_date,
        'endDate': end_date,
        'address': address,
        'organization': organization,
        'legalAddress': legalAddress,
        'CEO': CEO,
        'organizationINN': organization_inn,
        'organizationKPP': organization_kpp,
        'organizationORGN': organization_ogrn,
        'paymentAccountOrganization': paymentAccount_organization,
        'fullName': full_name,
        'passportSeries': passportSeries,
        'passportNumber': passportNumber,
        'temporaryRegistration': temporary_registration,
    }

    doc.render(context)

    global path_file
    path = path_file
    if os.path.exists(path + '/' + 'GPC_agreement.docx') == False:
        doc.save(path + '/' + 'GPC_agreement.docx')
    else:
        i = 1
        while True:
            if os.path.exists(path_file + '/' + f'GPC_agreement{i}.docx') == False:
                path = path_file + '/' + f'GPC_agreement{i}.docx'
                doc.save(path)
                break
            i += 1
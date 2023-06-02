import os
from docxtpl import DocxTemplate
from datetime import datetime
from document_generation_app.document_generation_functions.api import CompanyAPI, IndividualAPI
from document_generation_app.document_generation_functions.functions import Date_conversion, Get_path_file, \
    SurnameDeclension, FirstNameDeclension, LastNameDeclension
from django.http import FileResponse, HttpResponse, StreamingHttpResponse, HttpResponse
from django.core.files import File

path_file = Get_path_file()

def Generation_GPH_contract(validated_data):
    company = CompanyAPI()
    paymentAccount_organization = company['bank']['paymentAccount']
    organization = company["organizationalForm"] + ' "' + company["name"] + '"'
    city = company['legalAddress']["city"]
    legalAddress = company['legalAddress']['postalCode'] + ', ' + city + ' г, ' + company['legalAddress']["street"] + \
                   ', ' + company['legalAddress']["house"]
    organization_inn = company['inn']
    organization_kpp = company['kpp']
    organization_ogrn = company['ogrn']

    first_name_CEO = company['director']['fio']['firstName']
    surname_CEO = company['director']['fio']['secondName']
    patronymic_CEO = company['director']['fio']['patronymic']

    declension_first_name_CEO = FirstNameDeclension(first_name_CEO)
    declension_surname_CEO = SurnameDeclension(surname_CEO)
    declension_patronymic_CEO = LastNameDeclension(patronymic_CEO)

    if patronymic_CEO != None and patronymic_CEO != '' and patronymic_CEO != 'string':
        CEO = surname_CEO + ' ' + first_name_CEO + ' ' + patronymic_CEO
        CEO_declension = declension_surname_CEO + ' ' + declension_first_name_CEO + ' ' + declension_patronymic_CEO
    else:
        CEO = surname_CEO + ' ' + first_name_CEO
        CEO_declension = declension_surname_CEO + ' ' + declension_first_name_CEO

    individual = IndividualAPI()
    surname = individual['fio']['secondName']
    name = individual['fio']['firstName']
    patronymic = individual['fio']['patronymic']
    passportSeries = individual['passport']['serias']
    passportNumber = individual['passport']['number']
    tempRegAddress = 'г. ' + individual['tempRegAddress']['city'] + individual['tempRegAddress']['street'] + ', ' + \
        individual['tempRegAddress']['house']

    if patronymic != None and patronymic != '' and patronymic != 'string':
        full_name = surname + ' ' + name + ' ' + patronymic
    else:
        full_name = surname + ' ' + name

    number = validated_data['number']
    start_date = str(validated_data['start_date'].day) + '-' + str(validated_data['start_date'].month) + '-' + \
                 str(validated_data['start_date'].year)
    start_date = Date_conversion(start_date, 'word_month')
    end_date = str(validated_data['end_date'].day) + '-' + str(validated_data['end_date'].month) + '-' + \
                 str(validated_data['end_date'].year)
    end_date = Date_conversion(end_date, 'word_month')
    address = validated_data['address']
    # data table
    name_service = validated_data['name_service']
    price = validated_data['price']

    path_file_doc = 'document_generation_app/document_templates/GPC_agreement.docx'
    doc = DocxTemplate(path_file_doc)

    context = {
        'number': number,
        'startDate': start_date,
        'endDate': end_date,
        'address': address,
        'organization': organization,
        'city': city,
        'legalAddress': legalAddress,
        'ceoDeclension': CEO_declension,
        'CEO': CEO,
        'organizationINN': organization_inn,
        'organizationKPP': organization_kpp,
        'organizationORGN': organization_ogrn,
        'paymentAccountOrganization': paymentAccount_organization,
        'fullName': full_name,
        'passportSeries': passportSeries,
        'passportNumber': passportNumber,
        'temporaryRegistration': tempRegAddress,
        'nameService': name_service,
        'price': price,
    }

    doc.render(context)

    global path_file
    file = ''
    if os.path.exists(path_file + '/' + 'GPC_agreement.docx') == False:
        doc.save(path_file + '/' + 'GPC_agreement.docx')
        file = path_file + '/' + 'GPC_agreement.docx'
    else:
        i = 1
        while True:
            if os.path.exists(path_file + '/' + f'GPC_agreement{i}.docx') == False:
                path = path_file + '/' + f'GPC_agreement{i}.docx'
                doc.save(path)
                file = path_file + '/' + f'GPC_agreement{i}.docx'
                break
            i += 1

    filename = os.path.basename(file)
    response = HttpResponse(File(open(file, 'rb')), content_type='application/msword')
    response['Content-Length'] = os.path.getsize(file)
    response['Content-Disposition'] = "attachment; filename=%s" % filename
    return response
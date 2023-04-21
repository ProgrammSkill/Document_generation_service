import os
import requests
from django.shortcuts import render, redirect
from docxtpl import DocxTemplate
from win32com.shell import shell, shellcon
from datetime import datetime
import pytz
from document_generation_app.document_generation_functions.api import CompanyAPI, IndividualAPI
from document_generation_app.document_generation_functions.functions import Date_conversion, Get_path_file

path_file = Get_path_file()

def Generate_Generation_payment_order_for_advance_payment(request):
    if request.method == 'POST':
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

        number = request.POST.get('number')
        obj_start_date = datetime.strptime(request.POST.get('start_date'), '%Y-%m-%d')
        start_date = Date_conversion(obj_start_date.strftime("%d-%m-%Y"), 'word_month')
        obj_end_date = datetime.strptime(request.POST.get('end_date'), '%Y-%m-%d')
        end_date = Date_conversion(obj_end_date.strftime("%d-%m-%Y"), 'word_month')
        address = request.POST.get('address')
        table_doc = request.POST.get('table_doc')

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
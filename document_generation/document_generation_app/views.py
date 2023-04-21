import os
import pymssql as pymssql
import requests
from django.core.management.commands import shell
from django.http import HttpResponse
from django.shortcuts import render, redirect
from .forms import *
from docx import Document
from docxtpl import DocxTemplate
from docx.shared import *
from win32com.shell import shell, shellcon
from datetime import datetime
from openpyxl import *
import pytz

# import function of document generations
from document_generation_app.document_generation_functions.generation_about_arrival import Generation_about_arrival
from document_generation_app.document_generation_functions.generation_payment_order_for_advance_payment import Generate_Generation_payment_order_for_advance_payment
from document_generation_app.document_generation_functions.generation_gpc_contract import Generation_GPH_contract
from document_generation_app.document_generation_functions.generation_employment_contract import Generation_employment_contract_document


path_file = shell.SHGetKnownFolderPath(shellcon.FOLDERID_Downloads)



def CompanyAPI():
    # id_company = "b5deb54b-01bd-4bc3-87d0-21359b046e2a"
    # response = requests.get("http://secretochka.ru:48910/company-service/api/v1/manager/companyes/" + id_company).json()
    # actual_address = requests.get("http://secretochka.ru:48910/company-service/api/v1/companyes/" + id_company + "/info/actualaddresses").json()
    # response["ActualAddreses"] = actual_address
    response = {
            "id": "3fa85f64-5717-4562-b3fc-2c963f66afa6",
            "inn": "7804525530",
            "kpp": "780401001",
            "name": "Капитал Кадры",
            "organizationalForm": "ООО",
            "ogrn": "1117746358608",
            "okved": "123123123",
            "legalAddress": {
                "name": "195299, Г.Санкт-Петербург, пр-кт Гражданский, д. 119 ЛИТЕР А, офис 8"
            },
            "contactInfo": {
                "id": 0,
                "name": "Олег",
                "phone": "+78989321223",
                "email": "user@example.com"
            },
            "bank": {
                "id": 0,
                "bankId": "12341234",
                "correspondentAccount": "12345123451234512345",
                "paymentAccount": "12345123451234512345"
            },
            "ActualAddreses": [{"name": "195299, Г.Санкт-Петербург, пр-кт Гражданский, д. 119 ЛИТЕР А, офис 8"}]
    }

    return response

def IndividualAPI():
    pass

def Date_conversion(date, type=None):
    arr_date = date.split('-')
    day = arr_date[0]
    arr_month = ['января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря']
    month = arr_date[1]
    month = int(month)
    month_conversion = arr_month[month-1]
    year = arr_date[2]

    if type == 'quotes':
        if day[0] == '0':
            day = day[1]
        date = '\"' + day + '\" ' + month_conversion + ' ' + year + ' г.'
    elif type == 'word_month':
        if day[0] == '0':
            day = day[1]
        date = day + ' ' + month_conversion + ' ' + year + ' г.'
    else:
        date = day + '.' + arr_date[1] + '.' + year

    return date


# Create your views here.
def index(request):
    return render(request,'document_generation_app/index.html')

def Employment_contract_Document(request):
    return render(request,'document_generation_app/employment_contract.html')


def Generate_Employment_contract_Document(request):
    if request.method == 'POST':
        Generation_employment_contract_document(request)
    return redirect('employment_contract')


def GPC_Agreement(request):
    return render(request, 'document_generation_app/gpc_contract.html')


def Generate_GPC_Document(request):
    if request.method == 'POST':
        Generation_GPH_contract(request)
    return redirect('gpc_contract')



def Removal_order(request):
    return render(request, 'document_generation_app/removal_order.html')


def Generate_removal_older(request):
    if request.method =='POST':
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

    return redirect('removal_order')


def Notice_conclusion(request):
    return render(request, 'document_generation_app/notice_conclusion.html')


def Generate_Notice_conclusion(request):
    if request.method == 'POST':
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

        name = request.POST.get('name').upper()
        list_name = list(name)
        job_title = request.POST.get('job_title').upper()
        list_job = list(job_title)
        base = request.POST.get('base')
        obj_start_date = datetime.strptime(request.POST.get('start_date'), '%Y-%m-%d')
        start_date = obj_start_date.strftime("%d-%m-%Y")
        address = request.POST.get('address').upper()
        list_address = list(address)
        person = request.POST.get('person')

        path_file_doc = 'document_generation_app/document_templates/notice_conclusion.xlsx'
        doc = load_workbook(path_file_doc)
        sheet = doc.active

        list_columns = ['A', 'C', 'E', 'G', 'I', 'K', 'M', 'O', 'Q', 'S', 'U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI', 'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO', 'BQ']
        
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
        list_columns_phone = ['U', 'W', 'Y', 'AA', 'AC', 'AE', 'AG', 'AI', 'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK', 'BM', 'BO', 'BQ']
        for symbol in list_phone:
            for col in range(index, len(list_columns_phone)):
                cell = list_columns_phone[index] + str(row)
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


        if base == 'Трудовой договор':
            sheet['A167'] = 'X'
        elif base == 'Гражданско-правовой договор на выполнение работ (оказание услуг)':
            sheet['W167'] = 'X'


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

        if person == 'person_proxy':
            try:
                full_name = request.POST.get('full_name')
                passportSeries = request.POST.get('series')
                passportNumber = request.POST.get('number')
                obj_date_issue = datetime.strptime(request.POST.get('date_issue'), '%Y-%m-%d')
                date_issue = obj_date_issue.strftime("%d-%m-%Y")
                issued_by = request.POST.get('issued_by')

                sheet["AE201"] = full_name
                passportSeries = passportSeries[0] + passportSeries[1] + ' ' + passportSeries[2] + passportSeries[3]
                sheet["G203"] = passportSeries
                sheet["X203"] = passportNumber
                sheet["AR203"] = Date_conversion(date_issue)
                sheet["J205"] = issued_by
            except:
                pass

        else:
            sheet["AE201"] = CEO


        global path_file
        path = path_file
        if os.path.exists(path + '/' + 'notice_conclusion.xlsx') == False:
            doc.save(path + '/' + 'notice_conclusion.xlsx')
        else:
            i = 1
            while True:
                if os.path.exists(path_file + '/' + f'notice_conclusiont{i}.xlsx') == False:
                    path = path_file + '/' + f'notice_conclusiont{i}.xlsx'
                    doc.save(path)
                    break
                i += 1

    return redirect('notice_conclusion')


def Termination_noticen(request):
    return render(request, 'document_generation_app/termination_notice.html')


def Generate_Termination_notice(request):
    if request.method == 'POST':
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

        name = request.POST.get('name').upper()
        list_name = list(name)
        job_title = request.POST.get('job_title').upper()
        list_job = list(job_title)
        base = request.POST.get('base')
        obj_end_date = datetime.strptime(request.POST.get('end_date'), '%Y-%m-%d')
        end_date = obj_end_date.strftime("%d-%m-%Y")
        initiator = request.POST.get('initiator')
        person = request.POST.get('person')

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
                full_name = request.POST.get('full_name')
                passportSeries = request.POST.get('series')
                passportNumber = request.POST.get('number')
                obj_date_issue = datetime.strptime(request.POST.get('date_issue'), '%Y-%m-%d')
                date_issue = obj_date_issue.strftime("%d-%m-%Y")
                issued_by = request.POST.get('issued_by')

                sheet["AE193"] = full_name
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

    return redirect('termination_notice')


def Right_not_to_withhold_pit(request):
    return render(request, 'document_generation_app/right_not_to_withhold_pit.html')


def Generate_Right_not_to_withhold_pit(request):
    if request.method == 'POST':
        company = CompanyAPI()
        organization = company["organizationalForm"] + ' "' + company["name"] + '"'
        list_organization = list(organization.upper())
        phone = company["contactInfo"]["phone"].upper()
        list_phone = list(phone)
        inn = company['inn']
        list_inn = list(inn)
        kpp = company['kpp']
        ogrn = company['ogrn']
        list_ogrn = list(ogrn)
        paymentAccount = company['bank']['paymentAccount'].upper()
        correspondentAccount = company['bank']['correspondentAccount'].upper()
        legalAddress = company['legalAddress']["name"].upper()
        list_legalAddress = list(legalAddress)
        ActualAddreses = company['ActualAddreses'][0]["name"].upper()
        CEO = 'Сталюкова Екатерина Александровна'

        individual = 'Гайназаров Кайратбек'.upper()
        # passport_series = 'AC'.upper()
        # passport_number = '4348554'.upper()

        code = request.POST.get('code')
        list_code = list(code)


        path_file_doc = 'document_generation_app/document_templates/right_not_to_withhold_pit.xlsx'
        doc = load_workbook(path_file_doc)
        sheet = doc.active

        list_columns = ['X', 'AA', 'AD', 'AG', 'AJ',
                        'AM', 'AP', 'AS', 'AV', 'AY', 'BB', 'BH']

        row = 2
        index = 0
        for symbol in list_inn:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                break


        list_columns = ['X', 'AA', 'AD', 'AG', 'AJ',
                        'AM', 'AP', 'AS', 'AV']
        row = 4
        index = 0
        for symbol in list_ogrn:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                break


        # Input code
        sheet["CF8"] = list_code[0]
        sheet["CJ8"] = list_code[1]
        sheet["CO8"] = list_code[2]
        sheet["CT8"] = list_code[3]


        list_columns = ['B', 'C', 'E', 'F', 'G', 'I', 'L', 'N', 'P', 'S', 'U', 'W', 'Z', 'AC', 'AF', 'AI','AL',
                        'AO', 'AR', 'AU', 'AX', 'AY', 'BA', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BE', 'BJ', 'BK', 'BL',
                        'BM', 'BO', 'BP', 'BR', 'BS', 'BU', 'BV', 'BW', 'BX', 'BY', 'BZ', 'CA', 'CB', 'CC', 'CE']
        row = 11
        index = 0
        for symbol in list_organization:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                if list_columns[index] == 'CE':
                    sheet[f'{cell}'] = symbol
                    row += 2
                    index = 0
                    break
                sheet[f'{cell}'] = symbol
                index += 1
                break


        now_date = datetime.now(pytz.timezone('UTC'))
        now_date = Date_conversion(now_date.strftime('%d-%m-%Y'))
        list_now_date = list(now_date)

        list_columns = ['U', 'W', 'Z', 'AC', 'AF',
                        'AI', 'AL', 'AO', 'AR', 'AU']
        row = 46
        index = 0
        for symbol in list_now_date:
            for col in range(index, len(list_columns)):
                cell = list_columns[index] + str(row)
                sheet[f'{cell}'] = symbol
                index += 1
                break


        global path_file
        path = path_file
        if os.path.exists(path + '/' + 'right_not_to_withhold_pit.xlsx') == False:
            doc.save(path + '/' + 'right_not_to_withhold_pit.xlsx')
        else:
            i = 1
            while True:
                if os.path.exists(path_file + '/' + f'right_not_to_withhold_pit{i}.xlsx') == False:
                    path = path_file + '/' + f'right_not_to_withhold_pit{i}.xlsx'
                    doc.save(path)
                    break
                i += 1

    return redirect('right_not_to_withhold_pit')


def About_arrival(request):
    return render(request, 'document_generation_app/about_arrival.html')


def Generate_About_arrival(request):
    if request.method == 'POST':
        Generation_about_arrival(request)

    return redirect('about_arrival')

def Payment_order_for_advance_payment(request):
    return render(request, 'document_generation_app/payment_order_for_advance_payment.html')

def Generate_Payment_order_for_advance_payment(request):
    if request.method == 'POST':
        Generate_Generation_payment_order_for_advance_payment(request)

    return redirect('payment_order_for_advance_payment')


# {
#     "id": "1fa85f64-1727-4862-b3fc-2c963f66afa4",
#     "inn": "7804525530",
#     "kpp": "780401001",
#     "name": "Капитал Кадры",
#     "organizationalForm": "ООО",
#     "ogrn": "1117746358608",
#     "okved": "123456",
#     "legalAddress": {
#         "name": "195299, Г.Санкт-Петербург, пр-кт Гражданский, д. 119 ЛИТЕР А, офис 8"
#     },
#     "contactInfo": {
#         "id": 0,
#         "name": "Олег",
#         "phone": "+78989321223",
#         "email": "user@example.com"
#     },
#     "bank": {
#         "id": 0,
#         "bankId": "123456789",
#         "correspondentAccount": "12345123451234512345",
#         "paymentAccount": "12345123451234512345"
#     },
#     "ActualAddreses": [
#         {"name": "195299, Г.Санкт-Петербург, пр-кт Гражданский, д. 119 ЛИТЕР А, офис 8"}
#     ]
# }
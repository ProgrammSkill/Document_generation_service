import os
import requests
from django.shortcuts import render, redirect
from docxtpl import DocxTemplate
from win32com.shell import shell, shellcon
from datetime import datetime
import pytz
from document_generation_app.document_generation_functions.api import CompanyAPI, IndividualAPI
from document_generation_app.document_generation_functions.functions import Date_conversion, Get_path_file
from dateutil.relativedelta import relativedelta

path_file = Get_path_file()

dict_patent_cost = [
    {'region': 'Москва', 'price': 6600},
    {'region': 'Московская область', 'price': 6600},
    {'region': 'Санкт-Петербург', 'price': 4400},
    {'region': 'Ленинградская область', 'price': 4400},
    {'region': 'Республика Адыгея (Адыгея)', 'price': 4892},
    {'region': 'Республика Алтай', 'price': 3969},
    {'region': 'Республика Башкортостан', 'price': 4700},
    {'region': 'Республика Бурятия', 'price': 7763},
    {'region': 'Республика Дагестан', 'price': 4903},
    {'region': 'Республика Ингушетия', 'price': 4876},
    {'region': 'Кабардино-Балкарская Республика', 'price': 8172},
    {'region': 'Республика Калмыкия', 'price': 4086},
    {'region': 'Карачаево-Черкесская Республика', 'price': 5176},
    {'region': 'Республика Карелия', 'price': 8123},
    {'region': 'Республика Коми', 'price': 6320},
    {'region': 'Республика Крым', 'price': 4342},
    {'region': 'Республика Марий Эл', 'price': 5966},
    {'region': 'Республика Мордовия', 'price': 5170},
    # ==================================================
    {'region': 'Республика Саха (Якутия)', 'price': 12255},
    {'region': 'Республика Саха', 'price': 12255},
    {'region': 'Якутия', 'price': 12255},
    # ==================================================
    {'region': 'Республика Северная Осетия - Алания', 'price': 4331},
    # ==================================================
    {'region': 'Республика Татарстан (Татарстан)', 'price': 5666},
    {'region': 'Республика Татарстан', 'price': 5666},
    # ==================================================
    {'region': 'Республика Тыва', 'price': 5211},
    {'region': 'Удмуртская Республика', 'price': 5584},
    {'region': 'Республика Хакасия', 'price': 6701},
    {'region': 'Чеченская Республика', 'price': 2724},
    # ==================================================
    {'region': 'Чувашская Республика - Чувашия', 'price': 5448},
    {'region': 'Чувашская Республика', 'price': 5448},
    # ==================================================
    {'region': 'Алтайский край', 'price': 5312},
    {'region': 'Забайкальский край', 'price': 9725},
    {'region': 'Камчатский край', 'price': 7900},
    {'region': 'Краснодарский край', 'price': 6810},
    {'region': 'Красноярский край', 'price': 6701},
    {'region': 'Пермский край', 'price': 4700},
    {'region': 'Приморский край', 'price': 8717},
    {'region': 'Ставропольский край', 'price': 5993},
    {'region': 'Хабаровский край', 'price': 7355},
    {'region': 'Амурская область', 'price': 7104},
    {'region': 'Архангельская область', 'price': 4631},
    {'region': 'Астраханская область', 'price': 4974},
    {'region': 'Белгородская область', 'price': 5734},
    {'region': 'Брянская область', 'price': 5698},
    {'region': 'Владимирская область', 'price': 6350},
    {'region': 'Волгоградская область', 'price': 4903},
    {'region': 'Вологодская область', 'price': 5720},
    {'region': 'Воронежская область', 'price': 6020},
    {'region': 'Ивановская область', 'price': 4850},
    {'region': 'Иркутская область', 'price': 7984},
    {'region': 'Калининградская область', 'price': 6265},
    {'region': 'Калужская область', 'price': 6500},
    # ==================================================
    {'region': 'Кемеровская область - Кузбасс', 'price': 5857},
    {'region': 'Кемеровская область', 'price': 5857},
    {'region': 'Кузбасс', 'price': 5857},
    # ==================================================
    {'region': 'Кировская область', 'price': 5557},
    {'region': 'Костромская область', 'price': 4732},
    {'region': 'Курганская область', 'price': 5448},
    {'region': 'Курская область', 'price': 7273},
    {'region': 'Липецкая область', 'price': 5993},
    {'region': 'Магаданская область', 'price': 8172},
    {'region': 'Мурманская область', 'price': 7627},
    {'region': 'Нижегородская область', 'price': 6292},
    {'region': 'Новгородская область', 'price': 8036},
    {'region': 'Новосибирская область', 'price': 5474},
    {'region': 'Омская область', 'price': 4767},
    {'region': 'Оренбургская область', 'price': 5421},
    {'region': 'Орловская область', 'price': 5720},
    {'region': 'Пензенская область', 'price': 5312},
    {'region': 'Псковская область', 'price': 5835},
    {'region': 'Ростовская область', 'price': 4903},
    {'region': 'Рязанская область', 'price': 6674},
    {'region': 'Самарская область', 'price': 5720},
    {'region': 'Саратовская область', 'price': 5927},
    {'region': 'Сахалинская область', 'price': 7654},
    {'region': 'Свердловская область', 'price': 5950},
    {'region': 'Смоленская область', 'price': 5348},
    {'region': 'Тамбовская область', 'price': 4300},
    {'region': 'Тверская область', 'price': 8081},
    {'region': 'Томская область', 'price': 5448},
    {'region': 'Тульская область', 'price': 6280},
    {'region': 'Тюменская область', 'price': 7180},
    {'region': 'Ульяновская область', 'price': 4903},
    {'region': 'Челябинская область', 'price': 6200},
    {'region': 'Ярославская область', 'price': 5176},
    {'region': 'Севастополь', 'price': 6783},
    {'region': 'Еврейская автономная область', 'price': 6592},
    {'region': 'Ненецкий автономный округ', 'price': 6652},
    # ==================================================
    {'region': 'Ханты-Мансийский автономный округ - Югра', 'price': 6652},
    {'region': 'Ханты-Мансийский автономный округ', 'price': 6652},
    # ==================================================
    {'region': 'Чукотский автономный округ', 'price': 7900},
    {'region': 'Ямало-Ненецкий автономный округ', 'price': 10929}
]


def Generate_Generation_payment_order_for_advance_payment(request):
    if request.method == 'POST':

        company = CompanyAPI()
        organization = company["organizationalForm"] + ' "' + company["name"] + '"'
        organization_inn = company['inn']
        organization_kpp = company['kpp']
        ogrn_organization = company['ogrn']
        paymentAccount_organization = company['bank']['paymentAccount']
        correspondentAccount = company['bank']['correspondentAccount']

        individual = IndividualAPI()
        surname = individual['surname']
        name = individual['name']
        patronymic = individual['patronymic']

        if patronymic != None and patronymic != '':
            full_name = surname + ' ' + name + ' ' + patronymic
        else:
            full_name = surname + ' ' + name

        individual_inn = individual['inn']
        individual_kpp = individual['kpp']
        expiration_date = individual['patent']['expiration_date']
        expiration_date = datetime.strptime(expiration_date, '%Y-%m-%d')
        number_months = int(request.POST.get('number_months'))
        territory_of_action = individual['patent']['territory_of_action']

        sum = 0
        for obj in dict_patent_cost:
            print(obj)
            if obj['region'] == territory_of_action:
                sum = obj['price'] * number_months
                break

        sum = str(sum)
        if len(sum) == 4:
            sum = sum.replace(f'{sum[0]}', f'{sum[0]} ', 1)
        elif len(sum) == 5:
            sum = sum.replace(f'{sum[0] + sum[1]}', f'{sum[0] + sum[1]} ', 1)
        elif len(sum) == 6:
            sum = sum.replace(f'{sum[0] + sum[1] + sum[2]}', f'{sum[0] + sum[1] + sum[2]} ', 1)
        sum = sum + '-00'


        renewal_date = expiration_date + relativedelta(months=number_months)
        renewal_date = Date_conversion(renewal_date.strftime("%d-%m-%Y"))

        expiration_date = Date_conversion(expiration_date.strftime("%d-%m-%Y"))
        period = str(expiration_date) + ' - ' + str(renewal_date)

        now_date = datetime.now(pytz.timezone('UTC'))
        now_date = now_date.strftime('%d-%m-%Y')

        path_file_doc = 'document_generation_app/document_templates/generation_payment_order_for_advance_payment.docx'
        doc = DocxTemplate(path_file_doc)

        context = {
            'nowDate': Date_conversion(now_date),
            'organization': organization,
            'organizationINN': organization_inn,
            'organizationKPP': organization_kpp,
            'paymentAccountOrganization': paymentAccount_organization,
            'fullName': full_name,
            'individualINN': individual_inn,
            'individualKPP': individual_kpp,
            'sum': sum,
            'period': period,
            # 'kpp': kpp,
            # 'phone': phone,
            # 'legalAddress': legalAddress,
            # 'ActualAddreses': ActualAddreses,
            # 'paymentAccount': paymentAccount,
            # 'correspondentAccount': correspondentAccount,
            # 'CEO': CEO,
            'individual': individual
        }

        doc.render(context)

        global path_file
        path = path_file
        if os.path.exists(path + '/' + 'generation_payment_order_for_advance_payment.docx') == False:
            doc.save(path + '/' + 'generation_payment_order_for_advance_payment.docx')
        else:
            i = 1
            while True:
                if os.path.exists(path_file + '/' + f'generation_payment_order_for_advance_payment{i}.docx') == False:
                    path = path_file + '/' + f'generation_payment_order_for_advance_payment{i}.docx'
                    doc.save(path)
                    break
                i += 1
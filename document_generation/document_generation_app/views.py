import requests
from django.core.management.commands import shell
from django.http import HttpResponse
from django.shortcuts import render, redirect
from win32com.shell import shell, shellcon

# import function of document generations
from document_generation_app.document_generation_functions.generation_about_arrival import Generation_about_arrival
from document_generation_app.document_generation_functions.generation_payment_order_for_advance_payment import Generate_Generation_payment_order_for_advance_payment
from document_generation_app.document_generation_functions.generation_gpc_contract import Generation_GPH_contract
from document_generation_app.document_generation_functions.generation_employment_contract import Generation_employment_contract_document
from document_generation_app.document_generation_functions.generation_removal_older import Generation_removal_older
from document_generation_app.document_generation_functions.generation_notice_conclusion import Generation_notice_conclusion
from document_generation_app.document_generation_functions.generation_termination_notice import Generation_termination_notice
from document_generation_app.document_generation_functions.generation_right_not_to_withhold_pit import Generation_generation_right_not_to_withhold_pit


def index(request):
    return render(request, 'document_generation_app/index.html')

def Employment_contract_Document(request):
    return render(request, 'document_generation_app/employment_contract.html')


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
    if request.method == 'POST':
        Generation_removal_older(request)
    return redirect('removal_order')


def Notice_conclusion(request):
    return render(request, 'document_generation_app/notice_conclusion.html')


def Generate_Notice_conclusion(request):
    if request.method == 'POST':
        Generation_notice_conclusion(request)
    return redirect('notice_conclusion')


def Termination_noticen(request):
    return render(request, 'document_generation_app/termination_notice.html')


def Generate_Termination_notice(request):
    if request.method == 'POST':
        Generation_termination_notice(request)
    return redirect('termination_notice')


def Right_not_to_withhold_pit(request):
    return render(request, 'document_generation_app/right_not_to_withhold_pit.html')


def Generate_Right_not_to_withhold_pit(request):
    if request.method == 'POST':
        Generation_generation_right_not_to_withhold_pit(request)
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
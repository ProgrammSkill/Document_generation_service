import requests
from django.core.management.commands import shell
from django.http import HttpResponse
from django.shortcuts import render, redirect
from win32com.shell import shell, shellcon


def index(request):
    return render(request, 'document_generation_app/index.html')

def Employment_contract(request):
    return render(request, 'document_generation_app/employment_contract.html')


def GPC_Agreement(request):
    return render(request, 'document_generation_app/gpc_contract.html')


def Suspension_order(request):
    return render(request, 'document_generation_app/suspension_order.html')


def Notice_conclusion(request):
    return render(request, 'document_generation_app/notice_conclusion.html')


def Termination_noticen(request):
    return render(request, 'document_generation_app/termination_notice.html')


def Right_not_to_withhold_pit(request):
    return render(request, 'document_generation_app/right_not_to_withhold_pit.html')


def Arrival_notice(request):
    return render(request, 'document_generation_app/arrival_notice.html')


def Payment_order_for_advance_payment(request):
    return render(request, 'document_generation_app/payment_order_for_advance_payment.html')


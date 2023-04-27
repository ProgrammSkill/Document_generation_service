from django.urls import path
from document_generation_app.views import *

urlpatterns = [
    path("", index),
    path("employment_contract/", Employment_contract_Document, name='employment_contract'),
    path("gpc_contract/", GPC_Agreement, name='gpc_contract'),
    path("removal_order/", Removal_order, name='removal_order'),
    path("notice_conclusion/", Notice_conclusion, name='notice_conclusion'),
    path("termination_notice/", Termination_noticen, name='termination_notice'),
    path("right_not_to_withhold_pit/", Right_not_to_withhold_pit, name='right_not_to_withhold_pit'),
    path("about_arrival/", About_arrival, name='about_arrival'),
    path("payment_order_for_advance_payment/", Payment_order_for_advance_payment, name='payment_order_for_advance_payment'),
]
from django.urls import path
from document_generation_app.views import *

urlpatterns = [
    path("", index),
    path("employment_contract/", Employment_contract, name='employment_contract'),
    path("gpc_contract/", GPC_Agreement, name='gpc_contract'),
    path("suspension_order/", Suspension_order, name='suspension_order'),
    path("notice_conclusion/", Notice_conclusion, name='notice_conclusion'),
    path("termination_notice/", Termination_noticen, name='termination_notice'),
    path("right_not_to_withhold_pit/", Right_not_to_withhold_pit, name='right_not_to_withhold_pit'),
    path("arrival_notice/", Arrival_notice, name='arrival_notice'),
    path("payment_order_for_advance_payment/", Payment_order_for_advance_payment, name='payment_order_for_advance_payment'),
]
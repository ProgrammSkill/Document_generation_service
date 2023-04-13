from django.urls import path
from document_generation_app.views import *

urlpatterns = [
    path("", index),
    path("employment_contract/generation", Generate_Employment_contract_Document, name='employment_contract_generation'),
    path("employment_contract/", Employment_contract_Document, name='employment_contract'),
    path("gpc_contract/generation", Generate_GPC_Document, name='gpc_contract_generation'),
    path("gpc_contract/", GPC_Agreement, name='gpc_contract'),
    path("removal_order/generation", Generate_removal_older, name='removal_order_generation'),
    path("removal_order/", Removal_order, name='removal_order'),
    path("notice_conclusion/generation", Generate_Notice_conclusion, name='notice_conclusion_generation'),
    path("notice_conclusion/", Notice_conclusion, name='notice_conclusion'),
    path("termination_notice/generation", Generate_Termination_notice, name='termination_notice_generation'),
    path("termination_notice/", Termination_noticen, name='termination_notice'),
    path("right_not_to_withhold_pit/generation", Generate_Right_not_to_withhold_pit,
         name='right_not_to_withhold_pit_generation'),
    path("right_not_to_withhold_pit/", Right_not_to_withhold_pit, name='right_not_to_withhold_pit'),
]
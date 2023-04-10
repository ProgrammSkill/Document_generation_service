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
    path("notice_conclusion/", Notice_conclusion, name='notice_conclusion'),

]
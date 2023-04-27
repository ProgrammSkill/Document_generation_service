from django.urls import path

from . import views


urlpatterns = [
    path('employment_contract/', views.EmploymentContractAPIView.as_view()),
    path('gpc_contract/', views.GPCContractAPIView.as_view()),
    path('removal_older/', views.RemovalOrderAPIView.as_view()),
    path('payment_order_for_advance_payment/', views.GenerationPaymentOrderForAdvancePaymentIView.as_view())
]
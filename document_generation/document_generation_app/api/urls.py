from django.urls import path

from . import views


urlpatterns = [
    path('employment_contract/', views.EmploymentContractAPIView.as_view()),
    path('gpc_contract/', views.GPCContractAPIView.as_view()),
    path('suspension_order/', views.SuspensionOrderAPIView.as_view()),
    path('payment_order_for_advance_payment/', views.PaymentOrderForAdvancePaymentIView.as_view()),
    path('notice_conclusion/', views.NoticeConclusionAPIView.as_view()),
    path('termination_notice/', views.TerminationNoticeAPIView.as_view()),
    path('arrival_notice/', views.ArrivalNoticeAPIView.as_view()),
    path('right_not_to_withhold_pit/', views.RightNotToWithholdPitAPIView.as_view())
]
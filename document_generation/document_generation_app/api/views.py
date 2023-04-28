import requests
from django.core.management.commands import shell
from rest_framework.response import Response
from rest_framework.generics import CreateAPIView
from .serializers import SerializersEmploymentContract, SerializersGPCContract, SerializersSuspensionOrder,\
    SerializersGenerationPaymentOrderForAdvancePayment, SerializersNoticeConclusion, SerializersTerminationNotice,\
    SerializersArrivalNotice, SerializersRightNotToWithholdPit


class GPCContractAPIView(CreateAPIView):
    serializer_class = SerializersGPCContract

    def post(self, request):
        serializer = SerializersGPCContract(data=request.data)
        serializer.is_valid(raise_exception=True)
        serializer.save()
        return Response({'post': serializer.data})


class SuspensionOrderAPIView(CreateAPIView):
    serializer_class = SerializersSuspensionOrder

    def post(self, request):
        serializer = SerializersSuspensionOrder(data=request.data)
        serializer.is_valid(raise_exception=True)
        serializer.save()
        return Response({'post': serializer.data})


class PaymentOrderForAdvancePaymentIView(CreateAPIView):
    serializer_class = SerializersGenerationPaymentOrderForAdvancePayment

    def post(self, request):
        serializer = SerializersGenerationPaymentOrderForAdvancePayment(data=request.data)
        serializer.is_valid(raise_exception=True)
        serializer.save()
        return Response({'post': serializer.data})


class EmploymentContractAPIView(CreateAPIView):
    serializer_class = SerializersEmploymentContract

    def post(self, request):
        serializer = SerializersEmploymentContract(data=request.data)
        serializer.is_valid(raise_exception=True)
        serializer.save()
        return Response({'post': serializer.data})


class NoticeConclusionAPIView(CreateAPIView):
    serializer_class = SerializersNoticeConclusion

    def post(self, request):
        serializer = SerializersNoticeConclusion(data=request.data)
        serializer.is_valid(raise_exception=True)
        serializer.save()
        return Response({'post': serializer.data})


class TerminationNoticeAPIView(CreateAPIView):
    serializer_class = SerializersTerminationNotice

    def post(self, request):
        serializer = SerializersTerminationNotice(data=request.data)
        serializer.is_valid(raise_exception=True)
        serializer.save()
        return Response({'post': serializer.data})

class ArrivalNoticeAPIView(CreateAPIView):
    serializer_class = SerializersArrivalNotice

    def post(self, request):
        serializer = SerializersArrivalNotice(data=request.data)
        serializer.is_valid(raise_exception=True)
        serializer.save()
        return Response({'post': serializer.data})


class RightNotToWithholdPitAPIView(CreateAPIView):
    serializer_class = SerializersRightNotToWithholdPit

    def post(self, request):
        serializer = SerializersRightNotToWithholdPit(data=request.data)
        serializer.is_valid(raise_exception=True)
        serializer.save()
        return Response({'post': serializer.data})
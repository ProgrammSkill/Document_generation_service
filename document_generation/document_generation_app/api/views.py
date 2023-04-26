import requests
from django.core.management.commands import shell
from rest_framework.response import Response
from rest_framework.generics import CreateAPIView
from .serializers import SerializersEmploymentContract, SerializersGPCContract, SerializersRemovalOrder,\
    SerializersGenerationPaymentOrderForAdvancePayment


class GPCContractAPIView(CreateAPIView):
    serializer_class = SerializersGPCContract

    def post(self, request):
        serializer = SerializersGPCContract(data=request.data)
        serializer.is_valid(raise_exception=True)
        serializer.save()
        return Response({'post': serializer.data})
        # print({'post': request.data})
        # return Response({'post': request.data})


class RemovalOrderAPIView(CreateAPIView):
    serializer_class = SerializersRemovalOrder

    def post(self, request):
        return Response({'post': request.data})

class GenerationPaymentOrderForAdvancePaymentIView(CreateAPIView):
    serializer_class = SerializersGenerationPaymentOrderForAdvancePayment

    def post(self, request):
        print(request.data)
        return Response({'post': request.data})


class EmploymentContractAPIView(CreateAPIView):
    serializer_class = SerializersEmploymentContract
    def post(self, request):
        return Response({'post': request.data})

from rest_framework import serializers
from drf_spectacular.utils import extend_schema, OpenApiParameter, OpenApiExample
from drf_spectacular.types import OpenApiTypes
from document_generation_app.document_generation_functions.generation_gpc_contract import Generation_GPH_contract
from document_generation_app.document_generation_functions.generation_employment_contract import Generation_employment_contract_document

class SerializersGPCContract(serializers.Serializer):
    number = serializers.CharField(write_only=True, max_length=10)
    start_date = serializers.DateField(write_only=True)
    end_date = serializers.DateField(write_only=True)
    address = serializers.CharField(write_only=True, max_length=50)

    def create(self, validated_data):
        Generation_GPH_contract(validated_data)
        return validated_data


class SerializersRemovalOrder(serializers.Serializer):
    number = serializers.CharField(write_only=True, max_length=10)
    start_date = serializers.DateField(write_only=True)

    def create(self, validated_data):
        print(validated_data)
        return validated_data


class SerializersGenerationPaymentOrderForAdvancePayment(serializers.Serializer):
    number_months = serializers.IntegerField(write_only=True)

    def create(self, validated_data):
        print(validated_data)
        return validated_data

class SerializersEmploymentContract(serializers.Serializer):
    number = serializers.CharField(write_only=True, max_length=10)
    job_title = serializers.CharField(write_only=True, max_length=30)
    salary = serializers.IntegerField(write_only=True)
    urgent = serializers.CharField(write_only=True, max_length=50)
    start_date = serializers.DateField(write_only=True)
    end_date_urgent = serializers.DateField()
    cause = serializers.CharField()
    start_time = serializers.TimeField(write_only=True)
    ends_time = serializers.TimeField(write_only=True)

    def create(self, validated_data):
        Generation_employment_contract_document(validated_data)
        return validated_data

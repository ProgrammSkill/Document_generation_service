from rest_framework import serializers
from drf_spectacular.utils import extend_schema, OpenApiParameter, OpenApiExample
from drf_spectacular.types import OpenApiTypes

# import function of document generations
from document_generation_app.document_generation_functions.generation_arrival_notice import Generation_arrival_notice
from document_generation_app.document_generation_functions.generation_payment_order_for_advance_payment import Generation_payment_order_for_advance_payment
from document_generation_app.document_generation_functions.generation_gpc_contract import Generation_GPH_contract
from document_generation_app.document_generation_functions.generation_employment_contract import Generation_employment_contract_document
from document_generation_app.document_generation_functions.generation_suspension_order import Generation_suspension_order
from document_generation_app.document_generation_functions.generation_notice_conclusion import Generation_notice_conclusion
from document_generation_app.document_generation_functions.generation_termination_notice import Generation_termination_notice
from document_generation_app.document_generation_functions.generation_right_not_to_withhold_pit import Generation_generation_right_not_to_withhold_pit

class SerializersGPCContract(serializers.Serializer):
    number = serializers.CharField(write_only=True, max_length=10)
    start_date = serializers.DateField(write_only=True)
    end_date = serializers.DateField(write_only=True)
    address = serializers.CharField(write_only=True, max_length=50)

    def create(self, validated_data):
        Generation_GPH_contract(validated_data)
        return validated_data


class SerializersSuspensionOrder(serializers.Serializer):
    number = serializers.CharField(write_only=True, max_length=10)
    start_date = serializers.DateField(write_only=True)

    def create(self, validated_data):
        Generation_suspension_order(validated_data)
        return validated_data


class SerializersGenerationPaymentOrderForAdvancePayment(serializers.Serializer):
    number_months = serializers.IntegerField(write_only=True)

    def create(self, validated_data):
        Generation_payment_order_for_advance_payment(validated_data)
        return validated_data

class SerializersEmploymentContract(serializers.Serializer):
    number = serializers.CharField(write_only=True, max_length=10)
    job_title = serializers.CharField(write_only=True, max_length=30)
    salary = serializers.IntegerField(write_only=True)
    contract_type = serializers.CharField(write_only=True, max_length=50)
    start_date = serializers.DateField(write_only=True)
    end_date_urgent = serializers.DateField()
    start_time = serializers.TimeField(write_only=True)
    ends_time = serializers.TimeField(write_only=True)
    cause = serializers.CharField()

    def create(self, validated_data):
        Generation_employment_contract_document(validated_data)
        return validated_data

class SerializersNoticeConclusion(serializers.Serializer):
    name = serializers.CharField(write_only=True, max_length=100)
    job_title = serializers.CharField(write_only=True, max_length=100)
    base = serializers.CharField(write_only=True, max_length=100)
    start_date = serializers.DateField(write_only=True)
    address = serializers.CharField(max_length=100)
    person = serializers.CharField(write_only=True, max_length=50)
    full_name = serializers.CharField(write_only=True, max_length=55)
    series = serializers.CharField(max_length=4)
    number = serializers.CharField(max_length=8)
    date_issue = serializers.DateField(write_only=True)
    issued_by = serializers.CharField(write_only=True, max_length=100)

    def create(self, validated_data):
        Generation_notice_conclusion(validated_data)
        return validated_data

class SerializersTerminationNotice(serializers.Serializer):
    name = serializers.CharField(write_only=True, max_length=10)
    job_title = serializers.CharField(write_only=True, max_length=100)
    base = serializers.CharField(write_only=True, max_length=100)
    end_date = serializers.DateField(write_only=True)
    initiator = serializers.CharField(max_length=50)
    person = serializers.CharField(write_only=True, max_length=50)
    full_name = serializers.CharField(write_only=True, max_length=55)
    series = serializers.CharField(max_length=4)
    number = serializers.CharField(max_length=8)
    date_issue = serializers.DateField(write_only=True)
    issued_by = serializers.CharField(write_only=True, max_length=100)

    def create(self, validated_data):
        Generation_termination_notice(validated_data)
        return validated_data


class SerializersArrivalNotice(serializers.Serializer):
    document_type = serializers.CharField(write_only=True, max_length=50)
    series = serializers.CharField(max_length=4)
    number = serializers.CharField(max_length=8)
    date_issue = serializers.DateField(write_only=True)
    sell_by = serializers.DateField(write_only=True)
    purpose_departure = serializers.CharField(write_only=True, max_length=50)
    phone = serializers.CharField(max_length=20)
    job_title = serializers.CharField(write_only=True, max_length=100)
    end_date = serializers.DateField(write_only=True)
    receiving_side = serializers.CharField(max_length=50)

    surname_receiving_side = serializers.CharField(max_length=50)
    name_receiving_side = serializers.CharField(max_length=50)
    patronymic_receiving_side = serializers.CharField(max_length=50)
    type_of_identity_document = serializers.CharField(max_length=50)
    series_receiving_side = serializers.CharField(max_length=4)
    number_receiving_side = serializers.CharField(max_length=8)
    date_issue_receiving_side = serializers.DateField(write_only=True)
    sell_by_receiving_side = serializers.DateField(write_only=True)

    region = serializers.CharField(max_length=30)
    area = serializers.CharField(max_length=30)
    city = serializers.CharField(max_length=30)
    street = serializers.CharField(max_length=30)
    house = serializers.CharField(max_length=4)
    frame = serializers.CharField(max_length=5)
    structure = serializers.CharField(max_length=4)
    apartment = serializers.CharField(max_length=4)

    worker_place_region = serializers.CharField(write_only=True, max_length=30)
    worker_place_area = serializers.CharField(write_only=True, max_length=30)
    worker_place_city = serializers.CharField(write_only=True, max_length=30)
    worker_place_street = serializers.CharField(write_only=True, max_length=30)
    worker_place_house = serializers.CharField(write_only=True, max_length=4)
    worker_place_frame = serializers.CharField(max_length=5)
    worker_place_structure = serializers.CharField(max_length=4)
    worker_place_apartment = serializers.CharField(max_length=4)

    def create(self, validated_data):
        Generation_arrival_notice(validated_data)
        return validated_data


class SerializersRightNotToWithholdPit(serializers.Serializer):
    code = serializers.CharField(write_only=True, max_length=4)

    def create(self, validated_data):
        Generation_generation_right_not_to_withhold_pit(validated_data)
        return validated_data
from rest_framework import serializers
from drf_spectacular.utils import extend_schema, OpenApiParameter, OpenApiExample
from drf_spectacular.types import OpenApiTypes
from drf_yasg.utils import swagger_auto_schema

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
    name_service = serializers.CharField(write_only=True, max_length=50)
    price = serializers.IntegerField(write_only=True)

    def create(self, validated_data):
        Generation_GPH_contract(validated_data)
        return validated_data


class SerializersSuspensionOrder(serializers.Serializer):
    number = serializers.CharField(write_only=True, max_length=10)
    start_date = serializers.DateField(write_only=True)
    REASON_SUSPENSION_CHOICES = (
        ('действующего Вида на жительство', 'valid_residence_permit'),
        ('действующего патента', 'valid_patent'),
        ('действующего разрешения на временное проживание', 'valid_temporary_residence_permit'),
        ('справки о прохождении вакцинации от кори или отказа от вакцинации, заверенного врачом', 'getting_vaccinated'),
        ('справки о прохождении медосмотра', 'passing_medical_examination'),
        ('справки о сдаче анализа крови на антитела', 'passing_analysis'),
        ('чеков, подтверждающих авансовую оплату за патент', 'checks'),
    )
    reason_suspension = serializers.ChoiceField(choices=REASON_SUSPENSION_CHOICES)
    first_point_performer = serializers.CharField(write_only=True, max_length=50)
    second_point_performer = serializers.CharField(write_only=True, max_length=50)

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
    CONTRACT_TYPE_CHOICES = (
        ('Бессрочный договор', 'perpetual'),
        ('Срочный договор', 'urgent')
    )
    contract_type = serializers.ChoiceField(choices=CONTRACT_TYPE_CHOICES)
    start_date = serializers.DateField(write_only=True)
    end_date_urgent = serializers.DateField()
    start_time = serializers.TimeField(write_only=True)
    end_time = serializers.TimeField(write_only=True)
    cause = serializers.CharField()

    def create(self, validated_data):
        Generation_employment_contract_document(validated_data)
        return validated_data


class SerializersNoticeConclusion(serializers.Serializer):
    name_territorial_body = serializers.CharField(write_only=True, max_length=100)
    job_title = serializers.CharField(write_only=True, max_length=100)
    BASE_TYPE_CHOICES = (
        ('Трудовой договор', 'employment_contract'),
        ('Гражданско-правовой договор на выполнение работ (оказание услуг)', 'civil_contract')
    )
    base = serializers.ChoiceField(choices=BASE_TYPE_CHOICES)
    start_date = serializers.DateField(write_only=True)
    address = serializers.CharField(max_length=100)
    BASE_TYPE_PERSON = (
        ('Человек, который подаёт документы по доверенности', 'person_proxy'),
        ('Директор', 'director')
    )
    person = serializers.ChoiceField(choices=BASE_TYPE_PERSON)
    full_name = serializers.CharField(write_only=True, max_length=55, required=False)
    series = serializers.CharField(max_length=12, required=False)
    number = serializers.CharField(max_length=10,  required=False)
    date_issue = serializers.DateField(write_only=True)
    issued_by = serializers.CharField(write_only=True, max_length=100)

    def create(self, validated_data):
        Generation_notice_conclusion(validated_data)
        return validated_data

class SerializersTerminationNotice(serializers.Serializer):
    name_territorial_body = serializers.CharField(write_only=True, max_length=100)
    job_title = serializers.CharField(write_only=True, max_length=100)
    BASE_TYPE_CHOICES = (
        ('Трудовой договор', 'employment_contract'),
        ('Гражданско-правовой договор на выполнение работ (оказание услуг)', 'civil_contract')
    )
    base = serializers.ChoiceField(choices=BASE_TYPE_CHOICES)
    end_date = serializers.DateField(write_only=True)
    BASE_TYPE_INITIATOR = (
        ('Да', 'Yes'),
        ('Нет', 'no')
    )
    initiator = serializers.ChoiceField(choices=BASE_TYPE_INITIATOR)
    BASE_TYPE_PERSON = (
        ('Человек, который подаёт документы по доверенности', 'person_proxy'),
        ('Директор', 'director')
    )
    person = serializers.ChoiceField(choices=BASE_TYPE_PERSON)
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
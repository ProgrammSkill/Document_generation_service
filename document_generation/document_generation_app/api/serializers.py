from rest_framework import serializers


class TestSerializer(serializers.Serializer):
    field1 = serializers.CharField(write_only=True)
    field2 = serializers.FloatField(write_only=True)
    file = serializers.FileField(read_only=True)

    def create(self, validated_data):
        # TODO: Логика и надо возврашат словар
        return {"file": "<Файл>"}

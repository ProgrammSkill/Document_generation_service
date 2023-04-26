from rest_framework.generics import CreateAPIView

from .serializers import TestSerializer

class TestView(CreateAPIView):
    serializer_class = TestSerializer

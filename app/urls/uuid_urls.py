from django.urls import path
from ..views import (
    UUIDGeneratorView,
    UUIDValidateView,
    UUIDBulkView,
)

urlpatterns = [
    path("tools/uuid-generate", UUIDGeneratorView.as_view(), name="uuid-generate"),
    path("tools/uuid-validate", UUIDValidateView.as_view(), name="uuid-validate"),
    path("tools/uuid-bulk", UUIDBulkView.as_view(), name="uuid-bulk"),
]

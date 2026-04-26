from django.urls import path
from ..views import (
    TimestampConverterView,
    TimestampBatchView,
)

urlpatterns = [
    path("tools/timestamp", TimestampConverterView.as_view(), name="timestamp-converter"),
    path("tools/timestamp/batch", TimestampBatchView.as_view(), name="timestamp-batch"),
]

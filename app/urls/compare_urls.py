from django.urls import path
from ..views import (
    TextCompareView,
    FileCompareView,
)

urlpatterns = [
    path("tools/text-compare", TextCompareView.as_view(), name="text-compare"),
    path("tools/file-compare", FileCompareView.as_view(), name="file-compare"),
]

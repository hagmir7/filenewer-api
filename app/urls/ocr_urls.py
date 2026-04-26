from django.urls import path
from ..views import OCRPDFView

urlpatterns = [
    path("tools/ocr-pdf", OCRPDFView.as_view(), name="ocr-pdf"),
]

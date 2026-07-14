from django.urls import path
from ..views import OCRPDFView, YouTubeTranscriptView

urlpatterns = [
    path("tools/ocr-pdf", OCRPDFView.as_view(), name="ocr-pdf"),
    # path("tools/youtube-transcript", YouTubeTranscriptView.as_view(), name="youtube-transcript"),
]

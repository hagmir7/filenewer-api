from django.urls import path
from ..views import (
    Base64EncodeTextView,
    Base64EncodeFileView,
    Base64DecodeView,
    Base64ValidateView,
)

urlpatterns = [
    path("tools/base64-encode-text", Base64EncodeTextView.as_view(), name="base64-encode-text"),
    path("tools/base64-encode-file", Base64EncodeFileView.as_view(), name="base64-encode-file"),
    path("tools/base64-decode", Base64DecodeView.as_view(), name="base64-decode"),
    path("tools/base64-validate", Base64ValidateView.as_view(), name="base64-validate"),
]

from django.urls import path
from ..views import (
    FileEncryptView,
    FileDecryptView,
    FileHashView,
    FileChecksumView,
)

urlpatterns = [
    path("tools/file-encrypt", FileEncryptView.as_view(), name="file-encrypt"),
    path("tools/file-decrypt", FileDecryptView.as_view(), name="file-decrypt"),
    path("tools/file-hash", FileHashView.as_view(), name="file-hash"),
    path("tools/file-checksum", FileChecksumView.as_view(), name="file-checksum"),
]

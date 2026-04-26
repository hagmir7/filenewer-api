from django.urls import path
from ..views import (
    ExcelToCSVView,
    MarkdownToExcelView,
    ExcelToMarkdownView,
)

urlpatterns = [
    path("tools/excel-to-csv", ExcelToCSVView.as_view(), name="excel-to-csv"),
    path("tools/markdown-to-excel", MarkdownToExcelView.as_view(), name="markdown-to-excel"),
    path("tools/excel-to-markdown", ExcelToMarkdownView.as_view(), name="excel-to-markdown"),
]

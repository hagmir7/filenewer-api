from django.urls import path
from ..views import (
    ExcelToCSVView,
    MarkdownToExcelView,
    ExcelToMarkdownView,
    SQLFileToExcelView,
    SQLTextToExcelView,
)

urlpatterns = [
    path("tools/excel-to-csv", ExcelToCSVView.as_view(), name="excel-to-csv"),
    path("tools/markdown-to-excel", MarkdownToExcelView.as_view(), name="markdown-to-excel"),
    path("tools/excel-to-markdown", ExcelToMarkdownView.as_view(), name="excel-to-markdown"),
    path('tools/sql-file-to-excel', SQLFileToExcelView.as_view(), name='sql-file-to-excel'),
    path('tools/sql-text-to-excel', SQLTextToExcelView.as_view(), name='sql-text-to-excel'),
]

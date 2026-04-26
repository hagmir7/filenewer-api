from django.urls import path
from ..views import (
    CSVFileToSQLView,
    CSVTextToSQLView,
    CSVFileToJSONView,
    CSVTextToJSONView,
    CSVFileToExcelView,
    CSVTextToExcelView,
    CSVFileViewerView,
    CSVTextViewerView,
)

urlpatterns = [
    path("tools/csv-file-to-sql", CSVFileToSQLView.as_view(), name="csv-file-to-sql"),
    path("tools/csv-text-to-sql", CSVTextToSQLView.as_view(), name="csv-text-to-sql"),
    path("tools/csv-file-to-json", CSVFileToJSONView.as_view(), name="csv-file-to-json"),
    path("tools/csv-text-to-json", CSVTextToJSONView.as_view(), name="csv-text-to-json"),
    path("tools/csv-file-to-excel", CSVFileToExcelView.as_view(), name="csv-file-to-excel"),
    path("tools/csv-text-to-excel", CSVTextToExcelView.as_view(), name="csv-text-to-excel"),
    path("tools/csv-viewer-file", CSVFileViewerView.as_view(), name="csv-viewer-file"),
    path("tools/csv-viewer-text", CSVTextViewerView.as_view(), name="csv-viewer-text"),
]

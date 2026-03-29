from django.urls import path
from .views import *

urlpatterns = [
    path("tools/pdf-to-word", PDFToWordView.as_view(), name="pdf-to-word"),
    path("tools/csv-file-to-sql", CSVFileToSQLView.as_view(), name="csv-file-to-sql"),
    path("tools/csv-text-to-sql", CSVTextToSQLView.as_view(), name="csv-text-to-sql"),
    path("tools/csv-file-to-json", CSVFileToJSONView.as_view(), name="csv-file-to-json"),
    path("tools/csv-text-to-json", CSVTextToJSONView.as_view(), name="csv-text-to-json"),
    path('tools/csv-text-to-excel',         CSVTextToExcelView.as_view(),  name='csv-text-to-excel'),
    path('tools/csv-file-to-excel',    CSVFileToExcelView.as_view(),  name='csv-file-to-excel'),
    path('tools/json-file-to-csv',  JSONFileToCSVView.as_view(), name='json-file-to-csv'),
    path('tools/json-text-to-csv',  JSONTextToCSVView.as_view(), name='json-text-to-csv'),
    path('tools/excel-to-csv', ExcelToCSVView.as_view(), name='excel-to-csv'),
    path('tools/json-text-to-excel',       JSONTextToExcelView.as_view(), name='json-text-to-excel'),
    path('tools/json-file-to-excel',  JSONFileToExcelView.as_view(), name='json-file-to-excel'),
    path('tools/pdf-to-jpg',          PDFToJPGView.as_view(),        name='pdf-to-jpg'),
    path('tools/pdf-to-excel', PDFToExcelView.as_view(), name='pdf-to-excel'),
    path('tools/word-to-pdf',         WordToPDFView.as_view(),       name='word-to-pdf'),
    path("health/", HealthCheckView.as_view(), name="health-check"),
]

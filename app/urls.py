from django.urls import path
from .views import *

urlpatterns = [
    path("tools/pdf-to-word", PDFToWordView.as_view(), name="pdf-to-word"),
    path("tools/csv-file-to-sql", CSVFileToSQLView.as_view(), name="csv-file-to-sql"),
    path("tools/csv-text-to-sql", CSVTextToSQLView.as_view(), name="csv-text-to-sql"),
    path("tools/csv-file-to-json", CSVFileToJSONView.as_view(), name="csv-file-to-json"),
    path("tools/csv-text-to-json", CSVTextToJSONView.as_view(), name="csv-text-to-json"),
    path("health/", HealthCheckView.as_view(), name="health-check"),
]

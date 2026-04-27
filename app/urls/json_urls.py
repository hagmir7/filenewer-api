from django.urls import path
from ..views import (
    JSONFileToCSVView,
    JSONTextToCSVView,
    JSONFileToExcelView,
    JSONTextToExcelView,
    JSONFileFormatterView,
    JSONTextFormatterView,
    JSONFileToYAMLView,
    JSONTextToYAMLView,
    YAMLFileToJSONView,
    YAMLTextToJSONView,
)

urlpatterns = [
    path("tools/json-file-to-csv", JSONFileToCSVView.as_view(), name="json-file-to-csv"),
    path("tools/json-text-to-csv", JSONTextToCSVView.as_view(), name="json-text-to-csv"),
    path("tools/json-file-to-excel", JSONFileToExcelView.as_view(), name="json-file-to-excel"),
    path("tools/json-text-to-excel", JSONTextToExcelView.as_view(), name="json-text-to-excel"),
    path("tools/format-json-file", JSONFileFormatterView.as_view(), name="json-file-formatter"),
    path("tools/format-json-text", JSONTextFormatterView.as_view(), name="json-text-formatter"),
    path('tools/json-file-to-yaml', JSONFileToYAMLView.as_view(), name='json-file-to-yaml'),
    path('tools/json-text-to-yaml', JSONTextToYAMLView.as_view(), name='json-text-to-yaml'),
    path('tools/yaml-file-to-json', YAMLFileToJSONView.as_view(), name='yaml-file-to-json'),
    path('tools/yaml-text-to-json', YAMLTextToJSONView.as_view(), name='yaml-text-to-json'),
]

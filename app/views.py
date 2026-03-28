import logging
import os
from pathlib import Path

from django.http import HttpResponse
from rest_framework import status
from rest_framework.parsers import MultiPartParser
from rest_framework.response import Response
from rest_framework.views import APIView
from rest_framework.permissions import AllowAny
from .serializers import PDFUploadSerializer
from rest_framework.parsers import MultiPartParser, JSONParser
from .services import convert_pdf_to_docx

from .services import *

logger = logging.getLogger(__name__)


class PDFToWordView(APIView):
    permission_classes = [AllowAny]
    """
    POST /api/convert/
    ──────────────────
    Upload a PDF file and receive a Word (.docx) document in return.

    Request  : multipart/form-data  { file: <pdf> }
    Response : application/vnd.openxmlformats-officedocument.wordprocessingml.document
               (binary .docx download)

    Error responses are JSON:
        { "error": "<message>" }
    """

    parser_classes = [MultiPartParser]

    def post(self, request, *args, **kwargs):
        serializer = PDFUploadSerializer(data=request.data)

        if not serializer.is_valid():
            return Response(
                {"error": serializer.errors},
                status=status.HTTP_400_BAD_REQUEST,
            )

        uploaded_file = serializer.validated_data["file"]
        original_name = Path(uploaded_file.name).stem  # filename without extension

        logger.info(
            "Conversion request: file='%s', size=%d bytes, user=%s",
            uploaded_file.name,
            uploaded_file.size,
            request.META.get("REMOTE_ADDR", "unknown"),
        )

        try:
            pdf_bytes = uploaded_file.read()
            docx_bytes = convert_pdf_to_docx(pdf_bytes)
        except Exception as exc:
            logger.exception("Conversion failed for '%s'", uploaded_file.name)
            return Response(
                {"error": f"Conversion failed: {exc}"},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR,
            )

        # ── Stream the .docx back as a download ──────────────────────────
        response = HttpResponse(
            docx_bytes,
            content_type=(
                "application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.document"
            ),
        )
        response["Content-Disposition"] = f'attachment; filename="{original_name}.docx"'
        response["Content-Length"] = len(docx_bytes)
        logger.info(
            "Conversion successful: '%s.docx' (%d bytes)",
            original_name,
            len(docx_bytes),
        )
        return response


class HealthCheckView(APIView):
    """
    GET /api/health/
    ────────────────
    Returns 200 OK if the service is running.
    """

    def get(self, request, *args, **kwargs):
        return Response(
            {
                "status": "ok",
                "service": "pdf-to-word-api",
                "version": "1.0.0",
            }
        )


class CSVFileToSQLView(APIView):
    """
    POST /api/convert/file/
    Upload a .csv file → returns SQL as text or downloadable .sql file.

    Form fields:
        file       : CSV file (required)
        table_name : SQL table name (required)
        dialect    : sqlite | postgresql | mysql  (default: sqlite)
        output     : text | file                  (default: text)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        table_name = request.data.get("table_name", "").strip()
        dialect = request.data.get("dialect", "sqlite")
        output = request.data.get("output", "text")

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)
        if not file.name.endswith(".csv"):
            return Response({"error": "Only .csv files are accepted."}, status=400)
        if not table_name:
            return Response({"error": "table_name is required."}, status=400)
        if dialect not in ("sqlite", "postgresql", "mysql"):
            return Response(
                {"error": "dialect must be sqlite, postgresql, or mysql."}, status=400
            )

        # ── Convert ───────────────────────────────
        try:
            sql = csv_to_sql(file, table_name, dialect)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return text or file download ──────────
        if output == "file":
            response = HttpResponse(sql, content_type="text/plain")
            response["Content-Disposition"] = f'attachment; filename="{table_name}.sql"'
            return response

        return Response({"sql": sql, "table": table_name, "dialect": dialect})


class CSVTextToSQLView(APIView):
    """
    POST /api/convert/text/
    Send raw CSV text in JSON body → returns SQL.

    JSON body:
        {
            "csv"        : "col1,col2\\nval1,val2",
            "table_name" : "my_table",
            "dialect"    : "postgresql",
            "output"     : "text"
        }
    """

    parser_classes = [JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        csv_text = request.data.get("csv", "").strip()
        table_name = request.data.get("table_name", "").strip()
        dialect = request.data.get("dialect", "sqlite")
        output = request.data.get("output", "text")

        # ── Validate ──────────────────────────────
        if not csv_text:
            return Response({"error": "csv field is required."}, status=400)
        if not table_name:
            return Response({"error": "table_name is required."}, status=400)
        if dialect not in ("sqlite", "postgresql", "mysql"):
            return Response(
                {"error": "dialect must be sqlite, postgresql, or mysql."}, status=400
            )

        # ── Convert ───────────────────────────────
        try:
            sql = csv_to_sql(csv_text, table_name, dialect)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return text or file download ──────────
        if output == "file":
            response = HttpResponse(sql, content_type="text/plain")
            response["Content-Disposition"] = f'attachment; filename="{table_name}.sql"'
            return response

        return Response({"sql": sql, "table": table_name, "dialect": dialect})


from .services import csv_to_sql, csv_to_json  # ← add csv_to_json


class CSVFileToJSONView(APIView):
    """
    POST /api/convert/file/json/
    Upload a .csv file → returns JSON.

    Form fields:
        file      : CSV file (required)
        separator : column separator  (default: auto-detect)
        orient    : records | columns | values | index  (default: records)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        separator = request.data.get("separator", None)
        orient = request.data.get("orient", "records")

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)
        if not file.name.endswith(".csv"):
            return Response({"error": "Only .csv files are accepted."}, status=400)
        if orient not in ("records", "columns", "values", "index"):
            return Response(
                {"error": "orient must be records, columns, values, or index."},
                status=400,
            )

        # ── Convert ───────────────────────────────
        try:
            result = csv_to_json(file, separator=separator, orient=orient)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(
            {
                "data": result,
                "orient": orient,
                "count": (
                    len(result) if isinstance(result, list) else len(result.keys())
                ),
            }
        )


class CSVTextToJSONView(APIView):
    """
    POST /api/convert/text/json/
    Send raw CSV text in JSON body → returns JSON.

    JSON body:
        {
            "csv"      : "col1,col2\\nval1,val2",
            "separator": ";",
            "orient"   : "records"
        }
    """

    parser_classes = [JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        csv_text = request.data.get("csv", "").strip()
        separator = request.data.get("separator", None)
        orient = request.data.get("orient", "records")

        # ── Validate ──────────────────────────────
        if not csv_text:
            return Response({"error": "csv field is required."}, status=400)
        if orient not in ("records", "columns", "values", "index"):
            return Response(
                {"error": "orient must be records, columns, values, or index."},
                status=400,
            )

        # ── Convert ───────────────────────────────
        try:
            result = csv_to_json(csv_text, separator=separator, orient=orient)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(
            {
                "data": result,
                "orient": orient,
                "count": (
                    len(result) if isinstance(result, list) else len(result.keys())
                ),
            }
        )

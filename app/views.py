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


from .services import csv_to_sql, csv_to_json, json_to_csv  # ← add json_to_csv


class JSONFileToCSVView(APIView):
    """
    POST /api/convert/json/file/
    Upload a .json file → returns CSV.

    Form fields:
        file      : JSON file (required)
        separator : output column separator  (default: ,)
        output    : text | file              (default: text)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        separator = request.data.get("separator", ",")
        output = request.data.get("output", "text")

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)
        if not file.name.endswith(".json"):
            return Response({"error": "Only .json files are accepted."}, status=400)

        # ── Convert ───────────────────────────────
        try:
            csv_str = json_to_csv(file, separator=separator)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return text or file download ──────────
        if output == "file":
            filename = file.name.replace(".json", ".csv")
            response = HttpResponse(csv_str, content_type="text/csv")
            response["Content-Disposition"] = f'attachment; filename="{filename}"'
            return response

        return Response(
            {
                "csv": csv_str,
                "rows": csv_str.count("\n") - 1,
            }
        )


class JSONTextToCSVView(APIView):
    """
    POST /api/convert/json/text/
    Send raw JSON in body → returns CSV.

    JSON body:
        {
            "json"     : "[{...}, {...}]",
            "separator": ",",
            "output"   : "text"
        }
    """

    parser_classes = [JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        json_input = request.data.get("json")
        separator = request.data.get("separator", ",")
        output = request.data.get("output", "text")

        # ── Validate ──────────────────────────────
        if not json_input:
            return Response({"error": "json field is required."}, status=400)

        # ── Convert ───────────────────────────────
        try:
            csv_str = json_to_csv(json_input, separator=separator)
        except json.JSONDecodeError:
            return Response({"error": "Invalid JSON input."}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return text or file download ──────────
        if output == "file":
            response = HttpResponse(csv_str, content_type="text/csv")
            response["Content-Disposition"] = 'attachment; filename="output.csv"'
            return response

        return Response(
            {
                "csv": csv_str,
                "rows": csv_str.count("\n") - 1,
            }
        )


class ExcelToCSVView(APIView):
    """
    POST /api/convert/excel/
    Upload an Excel file (.xlsx/.xls) → returns CSV.

    Form fields:
        file       : Excel file               (required)
        sheet_name : sheet name or index      (default: all sheets)
        separator  : output separator         (default: ,)
        output     : text | file | zip        (default: text)
                     - text → JSON response with CSV strings
                     - file → download first sheet as .csv
                     - zip  → download all sheets as .zip of .csv files
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        sheet_name = request.data.get("sheet_name", None)
        separator = request.data.get("separator", ",")
        output = request.data.get("output", "text")

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)

        if not file.name.endswith((".xlsx", ".xls")):
            return Response(
                {"error": "Only .xlsx or .xls files are accepted."}, status=400
            )

        if output not in ("text", "file", "zip"):
            return Response({"error": "output must be text, file, or zip."}, status=400)

        # ── Parse sheet_name ──────────────────────
        if sheet_name is not None:
            try:
                sheet_name = int(sheet_name)  # try as index
            except (ValueError, TypeError):
                pass  # keep as string name

        # ── Convert ───────────────────────────────
        try:
            result = excel_to_csv(file, sheet_name=sheet_name, separator=separator)
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        sheets = result["sheets"]

        # ── Output: text (JSON response) ──────────
        if output == "text":
            return Response(
                {
                    "sheets": sheets,
                    "total_sheets": result["total_sheets"],
                    "sheet_names": result["sheet_names"],
                }
            )

        # ── Output: single .csv file ───────────────
        if output == "file":
            first_sheet_name = result["sheet_names"][0]
            csv_str = sheets[first_sheet_name]
            filename = file.name.rsplit(".", 1)[0] + ".csv"
            response = HttpResponse(csv_str, content_type="text/csv")
            response["Content-Disposition"] = f'attachment; filename="{filename}"'
            return response

        # ── Output: zip (all sheets as CSV files) ──
        if output == "zip":
            import zipfile

            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for name, csv_str in sheets.items():
                    safe_name = sanitize_name(name)
                    zf.writestr(f"{safe_name}.csv", csv_str)

            zip_buffer.seek(0)
            zip_name = file.name.rsplit(".", 1)[0] + ".zip"
            response = HttpResponse(zip_buffer.read(), content_type="application/zip")
            response["Content-Disposition"] = f'attachment; filename="{zip_name}"'
            return response







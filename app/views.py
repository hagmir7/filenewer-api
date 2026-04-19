import logging
import os
from pathlib import Path
import base64
import zipfile
from django.http import HttpResponse
from rest_framework import status
from rest_framework.parsers import MultiPartParser
from rest_framework.response import Response
from rest_framework.views import APIView
from rest_framework.permissions import AllowAny
from .serializers import PDFUploadSerializer
from rest_framework.parsers import MultiPartParser, JSONParser
from .services import convert_pdf_to_docx
from django.shortcuts import render
from .services import *

logger = logging.getLogger(__name__)


def index(request):
    return render(request, 'index.html')


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


class JSONFileToExcelView(APIView):
    """
    POST /api/convert/json/excel/file/
    Upload a .json file → returns .xlsx file.

    Form fields:
        file       : JSON file     (required)
        sheet_name : sheet name    (default: Sheet1)
        output     : text | file   (default: file)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        sheet_name = request.data.get("sheet_name", "Sheet1")
        output = request.data.get("output", "file")

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)
        if not file.name.endswith(".json"):
            return Response({"error": "Only .json files are accepted."}, status=400)

        # ── Convert ───────────────────────────────
        try:
            excel_bytes = json_to_excel(file, sheet_name=sheet_name)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return file download ───────────────────
        filename = file.name.replace(".json", ".xlsx")
        response = HttpResponse(
            excel_bytes,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response["Content-Disposition"] = f'attachment; filename="{filename}"'
        response["Content-Length"] = len(excel_bytes)
        return response


class JSONTextToExcelView(APIView):
    """
    POST /api/convert/json/excel/
    Send JSON in body → returns .xlsx file.

    JSON body (single sheet):
        {
            "json"      : [{...}, {...}],
            "sheet_name": "Users",
            "output"    : "file"
        }

    JSON body (multi-sheet):
        {
            "sheets": {
                "Users"  : [{...}, {...}],
                "Orders" : [{...}, {...}]
            }
        }
    """

    parser_classes = [JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        sheets_input = request.data.get("sheets")  # multi-sheet
        json_input = request.data.get("json")  # single sheet
        sheet_name = request.data.get("sheet_name", "Sheet1")
        output = request.data.get("output", "file")
        filename = request.data.get("filename", "output.xlsx")

        # ── Validate ──────────────────────────────
        if not sheets_input and not json_input:
            return Response(
                {
                    "error": 'Provide either "json" (single sheet) or "sheets" (multi-sheet).'
                },
                status=400,
            )

        if not filename.endswith(".xlsx"):
            filename += ".xlsx"

        # ── Convert ───────────────────────────────
        try:
            if sheets_input:
                # Multi-sheet mode
                if not isinstance(sheets_input, dict):
                    return Response(
                        {"error": '"sheets" must be an object: {"SheetName": [...]}'},
                        status=400,
                    )
                excel_bytes = json_to_excel_multisheets(sheets_input)
            else:
                # Single sheet mode
                excel_bytes = json_to_excel(json_input, sheet_name=sheet_name)

        except json.JSONDecodeError:
            return Response({"error": "Invalid JSON input."}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return file download ───────────────────
        response = HttpResponse(
            excel_bytes,
            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        response["Content-Disposition"] = f'attachment; filename="{filename}"'
        response["Content-Length"] = len(excel_bytes)
        return response


class CSVFileToExcelView(APIView):
    """
    POST /api/convert/csv/excel/file/
    Upload a .csv file → returns .xlsx file.

    Form fields:
        file       : CSV file           (required)
        sheet_name : sheet name         (default: Sheet1)
        separator  : column separator   (default: auto-detect)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        sheet_name = request.data.get("sheet_name", "Sheet1")
        separator = request.data.get("separator", None)

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)
        if not file.name.endswith(".csv"):
            return Response({"error": "Only .csv files are accepted."}, status=400)

        # ── Convert ───────────────────────────────
        try:
            excel_bytes = csv_to_excel(
                file,
                sheet_name=sheet_name,
                separator=separator,
            )
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return file download ───────────────────
        filename = file.name.replace(".csv", ".xlsx")
        response = HttpResponse(
            excel_bytes,
            content_type=(
                "application/vnd.openxmlformats-officedocument" ".spreadsheetml.sheet"
            ),
        )
        response["Content-Disposition"] = f'attachment; filename="{filename}"'
        response["Content-Length"] = len(excel_bytes)
        return response


class CSVTextToExcelView(APIView):
    """
    POST /api/convert/csv/excel/
    Send raw CSV text in JSON body → returns .xlsx file.

    JSON body (single sheet):
        {
            "csv"       : "col1,col2\\nval1,val2",
            "sheet_name": "Sheet1",
            "separator" : ",",
            "filename"  : "output.xlsx"
        }

    JSON body (multi-sheet):
        {
            "sheets": {
                "Users"  : "col1,col2\\nval1,val2",
                "Orders" : "col1,col2\\nval1,val2"
            },
            "separator": ",",
            "filename" : "report.xlsx"
        }
    """

    parser_classes = [JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        csv_input = request.data.get("csv")
        sheets_input = request.data.get("sheets")
        sheet_name = request.data.get("sheet_name", "Sheet1")
        separator = request.data.get("separator", None)
        filename = request.data.get("filename", "output.xlsx")

        # ── Validate ──────────────────────────────
        if not csv_input and not sheets_input:
            return Response(
                {
                    "error": 'Provide either "csv" (single sheet) or "sheets" (multi-sheet).'
                },
                status=400,
            )

        if not filename.endswith(".xlsx"):
            filename += ".xlsx"

        # ── Convert ───────────────────────────────
        try:
            if sheets_input:
                if not isinstance(sheets_input, dict):
                    return Response(
                        {
                            "error": '"sheets" must be an object: {"SheetName": "csv..."}'
                        },
                        status=400,
                    )
                excel_bytes = csv_to_excel_multisheets(
                    sheets_input, separator=separator
                )
            else:
                excel_bytes = csv_to_excel(
                    csv_input,
                    sheet_name=sheet_name,
                    separator=separator,
                )
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return file download ───────────────────
        response = HttpResponse(
            excel_bytes,
            content_type=(
                "application/vnd.openxmlformats-officedocument" ".spreadsheetml.sheet"
            ),
        )
        response["Content-Disposition"] = f'attachment; filename="{filename}"'
        response["Content-Length"] = len(excel_bytes)
        return response


class PDFToExcelView(APIView):
    """
    POST /api/convert/pdf/excel/
    Upload a PDF file → returns .xlsx file.

    Form fields:
        file     : PDF file              (required)
        password : PDF password          (optional)
        filename : output filename       (default: output.xlsx)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        password = request.data.get("password", None)
        filename = request.data.get("filename", None)

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)
        if not file.name.lower().endswith(".pdf"):
            return Response({"error": "Only .pdf files are accepted."}, status=400)

        if not filename:
            filename = file.name.replace(".pdf", ".xlsx").replace(".PDF", ".xlsx")
        if not filename.endswith(".xlsx"):
            filename += ".xlsx"

        # ── Convert ───────────────────────────────
        try:
            excel_bytes = pdf_to_excel(file, password=password)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return file download ───────────────────
        response = HttpResponse(
            excel_bytes,
            content_type=(
                "application/vnd.openxmlformats-officedocument" ".spreadsheetml.sheet"
            ),
        )
        response["Content-Disposition"] = f'attachment; filename="{filename}"'
        response["Content-Length"] = len(excel_bytes)
        return response


class PDFToJPGView(APIView):
    """
    POST /api/convert/pdf/jpg/
    Upload a PDF file → returns JPG image(s).

    Form fields:
        file     : PDF file                         (required)
        dpi      : image resolution                 (default: 200)
        quality  : JPG quality 1-95                 (default: 85)
        password : PDF password                     (optional)
        output   : single | zip | base64            (default: zip)
                   - single  → first page only as .jpg download
                   - zip     → all pages as .zip of .jpg files
                   - base64  → JSON with base64 encoded images
        page     : specific page number             (default: all)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        dpi = int(request.data.get("dpi", 200))
        quality = int(request.data.get("quality", 85))
        password = request.data.get("password", None)
        output = request.data.get("output", "zip")
        page = request.data.get("page", None)

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)
        if not file.name.lower().endswith(".pdf"):
            return Response({"error": "Only .pdf files are accepted."}, status=400)
        if output not in ("single", "zip", "base64"):
            return Response(
                {"error": "output must be single, zip, or base64."},
                status=400,
            )
        if not (1 <= quality <= 95):
            return Response(
                {"error": "quality must be between 1 and 95."},
                status=400,
            )
        if not (72 <= dpi <= 600):
            return Response(
                {"error": "dpi must be between 72 and 600."},
                status=400,
            )

        # ── Convert ───────────────────────────────
        try:
            pages = pdf_to_jpg(file, dpi=dpi, quality=quality, password=password)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        if not pages:
            return Response({"error": "No pages found in PDF."}, status=400)

        # ── Filter specific page if requested ──────
        if page is not None:
            try:
                page = int(page)
                pages = [p for p in pages if p["page"] == page]
                if not pages:
                    return Response(
                        {"error": f"Page {page} not found in PDF."},
                        status=400,
                    )
            except ValueError:
                return Response({"error": "page must be an integer."}, status=400)

        base_name = file.name.replace(".pdf", "").replace(".PDF", "")

        # ── Output: single page .jpg download ─────
        if output == "single":
            first = pages[0]
            response = HttpResponse(first["bytes"], content_type="image/jpeg")
            response["Content-Disposition"] = (
                f'attachment; filename="{base_name}_page_{first["page"]}.jpg"'
            )
            response["Content-Length"] = len(first["bytes"])
            return response

        # ── Output: zip of all pages ───────────────
        if output == "zip":
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for p in pages:
                    zf.writestr(
                        f'{base_name}_page_{p["page"]}.jpg',
                        p["bytes"],
                    )
            zip_buffer.seek(0)

            response = HttpResponse(
                zip_buffer.read(),
                content_type="application/zip",
            )
            response["Content-Disposition"] = (
                f'attachment; filename="{base_name}_pages.zip"'
            )
            return response

        # ── Output: base64 JSON response ───────────
        if output == "base64":
            data = [
                {
                    "page": p["page"],
                    "filename": f'{base_name}_page_{p["page"]}.jpg',
                    "width": p["width"],
                    "height": p["height"],
                    "base64": base64.b64encode(p["bytes"]).decode("utf-8"),
                    "size_kb": round(len(p["bytes"]) / 1024, 2),
                }
                for p in pages
            ]
            return Response(
                {
                    "total_pages": len(data),
                    "dpi": dpi,
                    "quality": quality,
                    "pages": data,
                }
            )


class WordToPDFView(APIView):
    """
    POST /api/convert/word/pdf/
    Upload a Word file (.docx / .doc) → returns PDF.

    Form fields:
        file     : Word file (.docx / .doc)   (required)
        filename : output filename             (default: document.pdf)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        filename = request.data.get("filename", None)

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)

        if not file.name.lower().endswith((".docx", ".doc")):
            return Response(
                {"error": "Only .docx or .doc files are accepted."},
                status=400,
            )

        if not filename:
            filename = (
                file.name.replace(".docx", ".pdf")
                .replace(".doc", ".pdf")
                .replace(".DOCX", ".pdf")
                .replace(".DOC", ".pdf")
            )

        if not filename.endswith(".pdf"):
            filename += ".pdf"

        # ── Convert ───────────────────────────────
        try:
            pdf_bytes = word_to_pdf(file, filename=file.name)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return file download ───────────────────
        response = HttpResponse(pdf_bytes, content_type="application/pdf")
        response["Content-Disposition"] = f'attachment; filename="{filename}"'
        response["Content-Length"] = len(pdf_bytes)
        return response


#

class EncryptPDFView(APIView):
    """
    POST /api/pdf/encrypt/
    Upload a PDF → returns encrypted PDF.

    Form fields:
        file             : PDF file              (required)
        user_password    : open password         (required)
        owner_password   : owner password        (default: same as user)
        allow_printing   : true | false          (default: true)
        allow_copying    : true | false          (default: true)
        allow_editing    : true | false          (default: true)
        allow_annotations: true | false          (default: true)
        filename         : output filename       (optional)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        user_password = request.data.get("user_password", "").strip()
        owner_password = request.data.get("owner_password", None)
        allow_printing = request.data.get("allow_printing", "true").lower() == "true"
        allow_copying = request.data.get("allow_copying", "true").lower() == "true"
        allow_editing = request.data.get("allow_editing", "true").lower() == "true"
        allow_annotations = (
            request.data.get("allow_annotations", "true").lower() == "true"
        )
        filename = request.data.get("filename", None)

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)
        if not file.name.lower().endswith(".pdf"):
            return Response({"error": "Only .pdf files are accepted."}, status=400)
        if not user_password:
            return Response({"error": "user_password is required."}, status=400)
        if len(user_password) < 4:
            return Response(
                {"error": "user_password must be at least 4 characters."},
                status=400,
            )

        if not filename:
            filename = file.name.replace(".pdf", "_encrypted.pdf")
        if not filename.endswith(".pdf"):
            filename += ".pdf"

        # ── Convert ───────────────────────────────
        try:
            encrypted_bytes = encrypt_pdf(
                file,
                user_password=user_password,
                owner_password=owner_password,
                allow_printing=allow_printing,
                allow_copying=allow_copying,
                allow_editing=allow_editing,
                allow_annotations=allow_annotations,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return encrypted PDF ───────────────────
        response = HttpResponse(encrypted_bytes, content_type="application/pdf")
        response["Content-Disposition"] = f'attachment; filename="{filename}"'
        response["Content-Length"] = len(encrypted_bytes)
        return response


class DecryptPDFView(APIView):
    """
    POST /api/pdf/decrypt/
    Upload an encrypted PDF + password → returns decrypted PDF.

    Form fields:
        file     : encrypted PDF file   (required)
        password : PDF password         (required)
        filename : output filename      (optional)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        password = request.data.get("password", "").strip()
        filename = request.data.get("filename", None)

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)
        if not file.name.lower().endswith(".pdf"):
            return Response({"error": "Only .pdf files are accepted."}, status=400)
        if not password:
            return Response({"error": "password is required."}, status=400)

        if not filename:
            filename = file.name.replace(".pdf", "_decrypted.pdf")
        if not filename.endswith(".pdf"):
            filename += ".pdf"

        # ── Decrypt ───────────────────────────────
        try:
            decrypted_bytes = decrypt_pdf(file, password=password)
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return decrypted PDF ───────────────────
        response = HttpResponse(decrypted_bytes, content_type="application/pdf")
        response["Content-Disposition"] = f'attachment; filename="{filename}"'
        response["Content-Length"] = len(decrypted_bytes)
        return response


class PDFInfoView(APIView):
    """
    POST /api/pdf/info/
    Upload a PDF → returns metadata and encryption info.

    Form fields:
        file     : PDF file     (required)
        password : PDF password (optional, for encrypted PDFs)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        password = request.data.get("password", None)

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)
        if not file.name.lower().endswith(".pdf"):
            return Response({"error": "Only .pdf files are accepted."}, status=400)

        # ── Get info ──────────────────────────────
        try:
            info = get_pdf_info(file, password=password)
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(info)


class CompressPDFView(APIView):
    """
    POST /api/pdf/compress/
    Upload a PDF → returns compressed PDF.

    Form fields:
        file              : PDF file                         (required)
        compression_level : low | medium | high | extreme   (default: medium)
        password          : PDF password                     (optional)
        output            : file | json                      (default: file)
        filename          : output filename                  (optional)

    output modes:
        file → download compressed PDF directly
        json → return compression stats + base64 PDF
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        compression_level = request.data.get("compression_level", "medium")
        password = request.data.get("password", None)
        output = request.data.get("output", "file")
        filename = request.data.get("filename", None)

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)

        if not file.name.lower().endswith(".pdf"):
            return Response({"error": "Only .pdf files are accepted."}, status=400)

        if compression_level not in ("low", "medium", "high", "extreme"):
            return Response(
                {"error": "compression_level must be: low, medium, high, or extreme."},
                status=400,
            )

        if output not in ("file", "json"):
            return Response(
                {"error": "output must be: file or json."},
                status=400,
            )

        # ── Output filename ───────────────────────
        if not filename:
            filename = file.name.replace(".pdf", f"_compressed_{compression_level}.pdf")
        if not filename.endswith(".pdf"):
            filename += ".pdf"

        # ── Compress ──────────────────────────────
        try:
            result = compress_pdf(
                file,
                compression_level=compression_level,
                password=password,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Output: JSON with stats + base64 ──────
        if output == "json":
            import base64

            return Response(
                {
                    "compression_level": result["compression_level"],
                    "original_size_kb": result["original_size_kb"],
                    "compressed_size_kb": result["compressed_size_kb"],
                    "original_size_mb": result["original_size_mb"],
                    "compressed_size_mb": result["compressed_size_mb"],
                    "reduction_percent": result["reduction_percent"],
                    "filename": filename,
                    "pdf_base64": base64.b64encode(result["bytes"]).decode("utf-8"),
                }
            )

        # ── Output: file download (default) ───────
        response = HttpResponse(
            result["bytes"],
            content_type="application/pdf",
        )
        response["Content-Disposition"] = f'attachment; filename="{filename}"'
        response["Content-Length"] = result["compressed_size"]
        response["X-Original-Size-KB"] = result["original_size_kb"]
        response["X-Compressed-Size-KB"] = result["compressed_size_kb"]
        response["X-Reduction-Percent"] = result["reduction_percent"]
        response["X-Compression-Level"] = result["compression_level"]
        return response


class PDFToPNGView(APIView):
    """
    POST /api/convert/pdf/png/
    Upload a PDF file → returns PNG image(s).

    Form fields:
        file     : PDF file                         (required)
        dpi      : image resolution 72-600          (default: 200)
        pages    : comma-separated page numbers     (default: all)
                   e.g. "1" or "1,3,5"
        password : PDF password                     (optional)
        output   : single | zip | base64            (default: zip)
                   - single  → first (or specific) page as .png
                   - zip     → all pages as .zip of .png files
                   - base64  → JSON with base64 encoded images
        filename : output filename                  (optional)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        dpi = request.data.get("dpi", "200")
        pages = request.data.get("pages", None)
        password = request.data.get("password", None)
        output = request.data.get("output", "zip")
        filename = request.data.get("filename", None)

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)

        if not file.name.lower().endswith(".pdf"):
            return Response({"error": "Only .pdf files are accepted."}, status=400)

        if output not in ("single", "zip", "base64"):
            return Response(
                {"error": "output must be: single, zip, or base64."},
                status=400,
            )

        # ── Parse DPI ─────────────────────────────
        try:
            dpi = int(dpi)
            if not (72 <= dpi <= 600):
                raise ValueError
        except ValueError:
            return Response(
                {"error": "dpi must be an integer between 72 and 600."},
                status=400,
            )

        # ── Parse pages ───────────────────────────
        parsed_pages = None
        if pages:
            try:
                parsed_pages = [int(p.strip()) for p in pages.split(",") if p.strip()]
            except ValueError:
                return Response(
                    {"error": 'pages must be comma-separated integers: e.g. "1,2,3"'},
                    status=400,
                )

        # ── Base filename ─────────────────────────
        base_name = file.name.replace(".pdf", "").replace(".PDF", "")

        # ── Convert ───────────────────────────────
        try:
            png_pages = pdf_to_png(
                file,
                dpi=dpi,
                pages=parsed_pages,
                password=password,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        if not png_pages:
            return Response({"error": "No pages found in PDF."}, status=400)

        # ── Output: single PNG ─────────────────────
        if output == "single":
            first = png_pages[0]
            out_name = filename or f'{base_name}_page_{first["page"]}.png'
            if not out_name.endswith(".png"):
                out_name += ".png"

            response = HttpResponse(first["bytes"], content_type="image/png")
            response["Content-Disposition"] = f'attachment; filename="{out_name}"'
            response["Content-Length"] = len(first["bytes"])
            response["X-Page"] = first["page"]
            response["X-Width"] = first["width"]
            response["X-Height"] = first["height"]
            return response

        # ── Output: ZIP of all PNGs ────────────────
        if output == "zip":
            import zipfile

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for p in png_pages:
                    zf.writestr(
                        f'{base_name}_page_{p["page"]}.png',
                        p["bytes"],
                    )
            zip_buffer.seek(0)

            out_name = filename or f"{base_name}_pages.zip"
            if not out_name.endswith(".zip"):
                out_name += ".zip"

            response = HttpResponse(
                zip_buffer.read(),
                content_type="application/zip",
            )
            response["Content-Disposition"] = f'attachment; filename="{out_name}"'
            response["X-Total-Pages"] = len(png_pages)
            return response

        # ── Output: base64 JSON ────────────────────
        if output == "base64":
            import base64

            data = [
                {
                    "page": p["page"],
                    "filename": f'{base_name}_page_{p["page"]}.png',
                    "width": p["width"],
                    "height": p["height"],
                    "size_kb": p["size_kb"],
                    "base64": base64.b64encode(p["bytes"]).decode("utf-8"),
                }
                for p in png_pages
            ]

            return Response(
                {
                    "total_pages": len(data),
                    "dpi": dpi,
                    "format": "PNG",
                    "pages": data,
                }
            )


class RotatePDFView(APIView):
    """
    POST /api/pdf/rotate/
    Upload a PDF → returns rotated PDF.

    Form fields:
        file     : PDF file                         (required)
        rotation : 90 | 180 | 270 | -90            (default: 90)
        pages    : comma-separated page numbers     (default: all)
                   e.g. "1" or "1,3,5"
        password : PDF password                     (optional)
        filename : output filename                  (optional)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        rotation = request.data.get("rotation", "90")
        pages = request.data.get("pages", None)
        password = request.data.get("password", None)
        filename = request.data.get("filename", None)

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)

        if not file.name.lower().endswith(".pdf"):
            return Response(
                {"error": "Only .pdf files are accepted."},
                status=400,
            )

        # ── Parse rotation ────────────────────────
        try:
            rotation = int(rotation)
        except ValueError:
            return Response(
                {"error": "rotation must be an integer: 90, 180, 270, or -90."},
                status=400,
            )

        # ── Parse pages ───────────────────────────
        parsed_pages = None
        if pages:
            try:
                parsed_pages = [int(p.strip()) for p in pages.split(",") if p.strip()]
                if not parsed_pages:
                    parsed_pages = None
            except ValueError:
                return Response(
                    {"error": 'pages must be comma-separated integers: e.g. "1,2,3"'},
                    status=400,
                )

        # ── Output filename ───────────────────────
        if not filename:
            filename = file.name.replace(".pdf", "_rotated.pdf")
        if not filename.endswith(".pdf"):
            filename += ".pdf"

        # ── Rotate ────────────────────────────────
        try:
            rotated_bytes, total_pages, rotated_count = rotate_pdf(
                file,
                rotation=rotation,
                pages=parsed_pages,
                password=password,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return rotated PDF ─────────────────────
        response = HttpResponse(rotated_bytes, content_type="application/pdf")
        response["Content-Disposition"] = f'attachment; filename="{filename}"'
        response["Content-Length"] = len(rotated_bytes)
        response["X-Total-Pages"] = total_pages
        response["X-Rotated-Pages"] = rotated_count
        response["X-Rotation-Degrees"] = rotation
        return response


class PDFPageInfoView(APIView):
    """
    POST /api/pdf/pages/
    Upload a PDF → returns page dimensions and current rotation per page.

    Form fields:
        file     : PDF file     (required)
        password : PDF password (optional)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        password = request.data.get("password", None)

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)

        if not file.name.lower().endswith(".pdf"):
            return Response(
                {"error": "Only .pdf files are accepted."},
                status=400,
            )

        # ── Get page info ─────────────────────────
        try:
            pages_info = get_pdf_page_info(file, password=password)
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(
            {
                "total_pages": len(pages_info),
                "pages": pages_info,
            }
        )


class WatermarkPDFView(APIView):
    """
    POST /api/pdf/watermark/
    Upload a PDF → returns watermarked PDF.

    Form fields (text watermark):
        file            : PDF file                              (required)
        watermark_text  : watermark text                       (default: CONFIDENTIAL)
        watermark_type  : text | image                         (default: text)
        opacity         : 0.0 - 1.0                            (default: 0.3)
        font_size       : text size                            (default: 60)
        color           : red|blue|grey|black|green|yellow|white (default: red)
        angle           : rotation angle                       (default: 45)
        position        : center|top|bottom|                   (default: center)
                          top-left|top-right|
                          bottom-left|bottom-right
        pages           : comma-separated page numbers         (default: all)
        password        : PDF password                         (optional)
        filename        : output filename                      (optional)

    Form fields (image watermark):
        file            : PDF file                             (required)
        watermark_type  : image                                (required)
        watermark_image : image file (png/jpg)                 (required)
        opacity         : 0.0 - 1.0                            (default: 0.3)
        angle           : rotation angle                       (default: 45)
        position        : position                             (default: center)
        pages           : comma-separated page numbers         (default: all)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        watermark_image = request.FILES.get("watermark_image", None)
        watermark_text = request.data.get("watermark_text", "CONFIDENTIAL")
        watermark_type = request.data.get("watermark_type", "text")
        opacity = request.data.get("opacity", "0.3")
        font_size = request.data.get("font_size", "60")
        color = request.data.get("color", "red")
        angle = request.data.get("angle", "45")
        position = request.data.get("position", "center")
        pages = request.data.get("pages", None)
        password = request.data.get("password", None)
        filename = request.data.get("filename", None)

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)

        if not file.name.lower().endswith(".pdf"):
            return Response(
                {"error": "Only .pdf files are accepted."},
                status=400,
            )

        if watermark_type == "image" and not watermark_image:
            return Response(
                {
                    "error": 'watermark_image is required when watermark_type is "image".'
                },
                status=400,
            )

        # ── Parse opacity ─────────────────────────
        try:
            opacity = float(opacity)
            if not (0.0 <= opacity <= 1.0):
                raise ValueError
        except ValueError:
            return Response(
                {"error": "opacity must be a float between 0.0 and 1.0."},
                status=400,
            )

        # ── Parse font_size ───────────────────────
        try:
            font_size = int(font_size)
            if not (8 <= font_size <= 200):
                raise ValueError
        except ValueError:
            return Response(
                {"error": "font_size must be an integer between 8 and 200."},
                status=400,
            )

        # ── Parse angle ───────────────────────────
        try:
            angle = int(angle)
        except ValueError:
            return Response(
                {"error": "angle must be an integer (e.g. 0, 45, 90)."},
                status=400,
            )

        # ── Parse pages ───────────────────────────
        parsed_pages = None
        if pages:
            try:
                parsed_pages = [int(p.strip()) for p in pages.split(",") if p.strip()]
            except ValueError:
                return Response(
                    {"error": 'pages must be comma-separated integers: e.g. "1,2,3"'},
                    status=400,
                )

        # ── Read watermark image bytes ─────────────
        wm_image_bytes = None
        if watermark_image:
            wm_image_bytes = watermark_image.read()

        # ── Output filename ───────────────────────
        if not filename:
            filename = file.name.replace(".pdf", "_watermarked.pdf")
        if not filename.endswith(".pdf"):
            filename += ".pdf"

        # ── Apply watermark ────────────────────────
        try:
            result_bytes = watermark_pdf(
                file,
                watermark_text=watermark_text,
                watermark_type=watermark_type,
                watermark_image=wm_image_bytes,
                opacity=opacity,
                font_size=font_size,
                color=color,
                angle=angle,
                position=position,
                pages=parsed_pages,
                password=password,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return watermarked PDF ─────────────────
        response = HttpResponse(result_bytes, content_type="application/pdf")
        response["Content-Disposition"] = f'attachment; filename="{filename}"'
        response["Content-Length"] = len(result_bytes)
        return response

class WatermarkPDFView(APIView):
    """
    POST /api/pdf/watermark/
    Upload a PDF → returns watermarked PDF.

    Form fields (text watermark):
        file            : PDF file                                (required)
        watermark_text  : watermark text                         (default: CONFIDENTIAL)
        watermark_type  : text | image                           (default: text)
        opacity         : 0.0 - 1.0                              (default: 0.3)
        font_size       : text size                              (default: 60)
        color           : red|blue|grey|black|green|yellow|white (default: red)
        angle           : rotation angle                         (default: 45)
        position        : center|top|bottom|                     (default: center)
                          top-left|top-right|
                          bottom-left|bottom-right
        pages           : comma-separated page numbers           (default: all)
        password        : PDF password                           (optional)
        filename        : output filename                        (optional)

    Form fields (image watermark):
        file            : PDF file                               (required)
        watermark_type  : image                                  (required)
        watermark_image : image file png/jpg                     (required)
        opacity         : 0.0 - 1.0                              (default: 0.3)
        angle           : rotation angle                         (default: 45)
        position        : position                               (default: center)
        pages           : comma-separated page numbers           (default: all)
        password        : PDF password                           (optional)
        filename        : output filename                        (optional)
    """
    parser_classes     = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file             = request.FILES.get('file')
        watermark_image  = request.FILES.get('watermark_image', None)
        watermark_text   = request.data.get('watermark_text',  'CONFIDENTIAL')
        watermark_type   = request.data.get('watermark_type',  'text')
        opacity          = request.data.get('opacity',         '0.3')
        font_size        = request.data.get('font_size',       '60')
        color            = request.data.get('color',           'red')
        angle            = request.data.get('angle',           '45')
        position         = request.data.get('position',        'center')
        pages            = request.data.get('pages',           None)
        password         = request.data.get('password',        None)
        filename         = request.data.get('filename',        None)

        # ── Validate file ──────────────────────────
        if not file:
            return Response(
                {'error': 'No file provided.'},
                status=400,
            )

        if not file.name.lower().endswith('.pdf'):
            return Response(
                {'error': 'Only .pdf files are accepted.'},
                status=400,
            )

        # ── Validate watermark type ────────────────
        if watermark_type not in ('text', 'image'):
            return Response(
                {'error': 'watermark_type must be "text" or "image".'},
                status=400,
            )

        if watermark_type == 'image' and not watermark_image:
            return Response(
                {'error': 'watermark_image file is required when watermark_type is "image".'},
                status=400,
            )

        if watermark_type == 'text' and not watermark_text.strip():
            return Response(
                {'error': 'watermark_text is required when watermark_type is "text".'},
                status=400,
            )

        # ── Validate opacity ───────────────────────
        try:
            opacity = float(opacity)
            if not (0.0 <= opacity <= 1.0):
                raise ValueError
        except ValueError:
            return Response(
                {'error': 'opacity must be a float between 0.0 and 1.0.'},
                status=400,
            )

        # ── Validate font_size ─────────────────────
        try:
            font_size = int(font_size)
            if not (8 <= font_size <= 200):
                raise ValueError
        except ValueError:
            return Response(
                {'error': 'font_size must be an integer between 8 and 200.'},
                status=400,
            )

        # ── Validate angle ─────────────────────────
        try:
            angle = int(angle)
        except ValueError:
            return Response(
                {'error': 'angle must be an integer e.g. 0, 45, 90.'},
                status=400,
            )

        # ── Validate color ─────────────────────────
        valid_colors = ('red', 'blue', 'grey', 'gray', 'black', 'green', 'yellow', 'white')
        if color.lower() not in valid_colors:
            return Response(
                {'error': f'color must be one of: {", ".join(valid_colors)}'},
                status=400,
            )

        # ── Validate position ──────────────────────
        valid_positions = (
            'center', 'top', 'bottom',
            'top-left', 'top-right',
            'bottom-left', 'bottom-right',
        )
        if position not in valid_positions:
            return Response(
                {'error': f'position must be one of: {", ".join(valid_positions)}'},
                status=400,
            )

        # ── Parse pages ───────────────────────────
        parsed_pages = None
        if pages:
            try:
                parsed_pages = [
                    int(p.strip())
                    for p in pages.split(',')
                    if p.strip()
                ]
                if not parsed_pages:
                    parsed_pages = None
            except ValueError:
                return Response(
                    {'error': 'pages must be comma-separated integers e.g. "1,2,3"'},
                    status=400,
                )

        # ── Read watermark image bytes ─────────────
        wm_image_bytes = None
        if watermark_image:
            try:
                wm_image_bytes = watermark_image.read()
                if not wm_image_bytes:
                    return Response(
                        {'error': 'watermark_image file is empty.'},
                        status=400,
                    )
            except Exception as e:
                return Response(
                    {'error': f'Failed to read watermark image: {e}'},
                    status=400,
                )

        # ── Output filename ───────────────────────
        if not filename:
            filename = file.name.replace('.pdf', '_watermarked.pdf')
            filename = filename.replace('.PDF', '_watermarked.pdf')
        if not filename.endswith('.pdf'):
            filename += '.pdf'

        # ── Apply watermark ────────────────────────
        try:
            result_bytes = watermark_pdf(
                source          = file,
                watermark_text  = watermark_text,
                watermark_type  = watermark_type,
                watermark_image = wm_image_bytes,
                opacity         = opacity,
                font_size       = font_size,
                color           = color,
                angle           = angle,
                position        = position,
                pages           = parsed_pages,
                password        = password,
            )
        except ValueError as e:
            return Response({'error': str(e)}, status=400)
        except RuntimeError as e:
            return Response({'error': str(e)}, status=500)
        except Exception as e:
            return Response({'error': f'Unexpected error: {e}'}, status=500)

        # ── Safety check ───────────────────────────
        if not result_bytes:
            return Response(
                {'error': 'Watermark failed — no output generated.'},
                status=500,
            )

        # ── Return watermarked PDF ─────────────────
        response = HttpResponse(result_bytes, content_type='application/pdf')
        response['Content-Disposition'] = f'attachment; filename="{filename}"'
        response['Content-Length']       = len(result_bytes)
        return response


class JSONFileFormatterView(APIView):
    """
    POST /api/format/json/file/
    Upload a .json file → returns formatted JSON.

    Form fields:
        file         : JSON file              (required)
        indent       : spaces for indent      (default: 4)
        sort_keys    : true | false           (default: false)
        minify       : true | false           (default: false)
        ensure_ascii : true | false           (default: false)
        output       : text | file            (default: text)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        indent = request.data.get("indent", "4")
        sort_keys = request.data.get("sort_keys", "false").lower() == "true"
        minify = request.data.get("minify", "false").lower() == "true"
        ensure_ascii = request.data.get("ensure_ascii", "false").lower() == "true"
        output = request.data.get("output", "text")

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)
        if not file.name.lower().endswith(".json"):
            return Response({"error": "Only .json files are accepted."}, status=400)
        if output not in ("text", "file"):
            return Response({"error": "output must be text or file."}, status=400)

        # ── Parse indent ──────────────────────────
        try:
            indent = int(indent)
            if not (0 <= indent <= 8):
                raise ValueError
        except ValueError:
            return Response(
                {"error": "indent must be an integer between 0 and 8."},
                status=400,
            )

        # ── Format ────────────────────────────────
        try:
            result = format_json(
                file,
                indent=indent,
                sort_keys=sort_keys,
                minify=minify,
                ensure_ascii=ensure_ascii,
            )
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Handle invalid JSON ────────────────────
        if not result["is_valid"]:
            return Response(
                {
                    "is_valid": False,
                    "error": result.get("error"),
                    "error_line": result.get("error_line"),
                    "error_column": result.get("error_column"),
                    "error_position": result.get("error_position"),
                    "size_original": result.get("size_original"),
                },
                status=422,
            )

        # ── Return file download ───────────────────
        if output == "file":
            response = HttpResponse(
                result["formatted"],
                content_type="application/json",
            )
            filename = file.name.replace(".json", "_formatted.json")
            response["Content-Disposition"] = f'attachment; filename="{filename}"'
            return response

        # ── Return JSON response ───────────────────
        return Response(
            {
                "is_valid": result["is_valid"],
                "formatted": result["formatted"],
                "type": result["type"],
                "depth": result["depth"],
                "key_count": result["key_count"],
                "item_count": result["item_count"],
                "size_original_kb": result["size_original_kb"],
                "size_formatted_kb": result["size_formatted_kb"],
                "minified": result["minified"],
                "sorted_keys": result["sorted_keys"],
                "indent": result["indent"],
            }
        )


class JSONTextFormatterView(APIView):
    """
    POST /api/format/json/text/
    Send raw JSON string → returns formatted JSON.

    JSON body:
        {
            "json"        : "{...}",
            "indent"      : 4,
            "sort_keys"   : false,
            "minify"      : false,
            "ensure_ascii": false,
            "output"      : "text"
        }
    """

    parser_classes = [JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        json_input = request.data.get("json")
        indent = request.data.get("indent", 4)
        sort_keys = request.data.get("sort_keys", False)
        minify = request.data.get("minify", False)
        ensure_ascii = request.data.get("ensure_ascii", False)
        output = request.data.get("output", "text")

        # ── Validate ──────────────────────────────
        if json_input is None:
            return Response({"error": '"json" field is required.'}, status=400)

        # Accept object/array directly or as string
        if isinstance(json_input, (dict, list)):
            import json as _json

            json_input = _json.dumps(json_input)

        if not isinstance(json_input, str):
            return Response(
                {"error": '"json" must be a string, object, or array.'},
                status=400,
            )

        if output not in ("text", "file"):
            return Response(
                {"error": "output must be text or file."},
                status=400,
            )

        # ── Validate indent ───────────────────────
        try:
            indent = int(indent)
            if not (0 <= indent <= 8):
                raise ValueError
        except (ValueError, TypeError):
            return Response(
                {"error": "indent must be an integer between 0 and 8."},
                status=400,
            )

        # ── Format ────────────────────────────────
        try:
            result = format_json(
                json_input,
                indent=indent,
                sort_keys=bool(sort_keys),
                minify=bool(minify),
                ensure_ascii=bool(ensure_ascii),
            )
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Handle invalid JSON ────────────────────
        if not result["is_valid"]:
            return Response(
                {
                    "is_valid": False,
                    "error": result.get("error"),
                    "error_line": result.get("error_line"),
                    "error_column": result.get("error_column"),
                    "error_position": result.get("error_position"),
                    "size_original": result.get("size_original"),
                },
                status=422,
            )

        # ── Return file download ───────────────────
        if output == "file":
            response = HttpResponse(
                result["formatted"],
                content_type="application/json",
            )
            response["Content-Disposition"] = 'attachment; filename="formatted.json"'
            return response

        # ── Return JSON response ───────────────────
        return Response(
            {
                "is_valid": result["is_valid"],
                "formatted": result["formatted"],
                "type": result["type"],
                "depth": result["depth"],
                "key_count": result["key_count"],
                "item_count": result["item_count"],
                "size_original_kb": result["size_original_kb"],
                "size_formatted_kb": result["size_formatted_kb"],
                "minified": result["minified"],
                "sorted_keys": result["sorted_keys"],
                "indent": result["indent"],
            }
        )

class OCRPDFView(APIView):
    """
    POST /api/pdf/ocr/
    Upload a PDF → extracts text using OCR.

    Form fields:
        file          : PDF file                          (required)
        language      : tesseract language code          (default: eng)
                        eng | ara | eng+ara | fra | deu
                        spa | chi_sim | rus | jpn | por
        dpi           : render resolution 150-600        (default: 300)
        pages         : comma-separated page numbers     (default: all)
        password      : PDF password                     (optional)
        output        : json | text | file               (default: json)
                        json → structured JSON response
                        text → plain text response
                        file → download .txt file
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        language = request.data.get("language", "eng")
        dpi = request.data.get("dpi", "300")
        pages = request.data.get("pages", None)
        password = request.data.get("password", None)
        output = request.data.get("output", "json")

        # ── Validate ──────────────────────────────
        if not file:
            return Response({"error": "No file provided."}, status=400)

        if not file.name.lower().endswith(".pdf"):
            return Response(
                {"error": "Only .pdf files are accepted."},
                status=400,
            )

        if output not in ("json", "text", "file"):
            return Response(
                {"error": "output must be: json, text, or file."},
                status=400,
            )

        # ── Validate DPI ──────────────────────────
        try:
            dpi = int(dpi)
            if not (72 <= dpi <= 600):
                raise ValueError
        except ValueError:
            return Response(
                {"error": "dpi must be an integer between 72 and 600."},
                status=400,
            )

        # ── Validate language ─────────────────────
        if not language.strip():
            return Response(
                {"error": "language is required."},
                status=400,
            )

        # ── Parse pages ───────────────────────────
        parsed_pages = None
        if pages:
            try:
                parsed_pages = [int(p.strip()) for p in pages.split(",") if p.strip()]
            except ValueError:
                return Response(
                    {"error": 'pages must be comma-separated integers: e.g. "1,2,3"'},
                    status=400,
                )

        # ── Run OCR ───────────────────────────────
        try:
            result = ocr_pdf(
                file,
                language=language,
                dpi=dpi,
                pages=parsed_pages,
                password=password,
            )
        except RuntimeError as e:
            return Response({"error": str(e)}, status=503)
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        base_name = file.name.replace(".pdf", "").replace(".PDF", "")

        # ── Output: plain text response ───────────
        if output == "text":
            return Response(
                {
                    "text": result["full_text"],
                    "total_pages": result["total_pages"],
                    "word_count": result["word_count"],
                    "char_count": result["char_count"],
                    "language": result["language"],
                    "dpi": result["dpi"],
                }
            )

        # ── Output: download .txt file ────────────
        if output == "file":
            response = HttpResponse(
                result["full_text"],
                content_type="text/plain; charset=utf-8",
            )
            response["Content-Disposition"] = (
                f'attachment; filename="{base_name}_ocr.txt"'
            )
            response["Content-Length"] = len(result["full_text"].encode("utf-8"))
            return response

        # ── Output: full JSON (default) ────────────
        return Response(
            {
                "full_text": result["full_text"],
                "pages": result["pages"],
                "total_pages": result["total_pages"],
                "word_count": result["word_count"],
                "char_count": result["char_count"],
                "language": result["language"],
                "dpi": result["dpi"],
            }
        )


class TimestampConverterView(APIView):
    """
    POST /api/tools/timestamp/
    Convert a single timestamp to multiple formats.

    JSON body:
        {
            "value"      : "1704067200",
            "from_tz"    : "UTC",
            "to_tz"      : "Asia/Riyadh",
            "from_format": null,
            "to_format"  : "%d/%m/%Y"
        }

    value accepts:
        - Unix seconds  : "1704067200"
        - Unix ms       : "1704067200000"
        - ISO 8601      : "2024-01-01T00:00:00Z"
        - Human         : "2024-01-01 12:00:00"
        - Date only     : "2024-01-01"
        - Relative      : "now" | "today" | "yesterday" | "tomorrow"
    """

    parser_classes = [JSONParser, MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        value = request.data.get("value")
        from_tz = request.data.get("from_tz", "UTC")
        to_tz = request.data.get("to_tz", "UTC")
        from_format = request.data.get("from_format", None)
        to_format = request.data.get("to_format", None)

        # ── Validate ──────────────────────────────
        if value is None:
            return Response(
                {"error": '"value" field is required.'},
                status=400,
            )

        # ── Convert ───────────────────────────────
        try:
            result = convert_timestamp(
                str(value),
                from_tz=from_tz,
                to_tz=to_tz,
                from_format=from_format,
                to_format=to_format,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(result)


class TimestampBatchView(APIView):
    """
    POST /api/tools/timestamp/batch/
    Convert multiple timestamps at once.

    JSON body:
        {
            "values"   : ["1704067200", "1704153600", "now"],
            "from_tz"  : "UTC",
            "to_tz"    : "Asia/Riyadh",
            "to_format": "%Y-%m-%d %H:%M:%S"
        }
    """

    parser_classes = [JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        values = request.data.get("values")
        from_tz = request.data.get("from_tz", "UTC")
        to_tz = request.data.get("to_tz", "UTC")
        to_format = request.data.get("to_format", None)

        # ── Validate ──────────────────────────────
        if not values:
            return Response(
                {"error": '"values" list is required.'},
                status=400,
            )
        if not isinstance(values, list):
            return Response(
                {"error": '"values" must be a list.'},
                status=400,
            )
        if len(values) > 100:
            return Response(
                {"error": "Maximum 100 timestamps per batch."},
                status=400,
            )

        # ── Convert ───────────────────────────────
        try:
            results = batch_convert_timestamps(
                values,
                from_tz=from_tz,
                to_tz=to_tz,
                to_format=to_format,
            )
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        success = sum(1 for r in results if r.get("success"))
        failed = len(results) - success

        return Response(
            {
                "total": len(results),
                "success": success,
                "failed": failed,
                "results": results,
            }
        )


class Base64EncodeTextView(APIView):
    """
    POST /api/tools/base64/encode/text/
    Encode raw text → Base64.

    JSON body:
        {
            "text"      : "Hello, World!",
            "encoding"  : "standard",
            "chunk_size": 0
        }
    """

    parser_classes = [JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        text = request.data.get("text")
        encoding = request.data.get("encoding", "standard")
        chunk_size = request.data.get("chunk_size", 0)

        # ── Validate ──────────────────────────────
        if text is None:
            return Response(
                {"error": '"text" field is required.'},
                status=400,
            )
        if encoding not in ("standard", "url_safe", "mime"):
            return Response(
                {"error": "encoding must be: standard, url_safe, or mime."},
                status=400,
            )
        try:
            chunk_size = int(chunk_size)
        except (ValueError, TypeError):
            return Response(
                {"error": "chunk_size must be an integer."},
                status=400,
            )

        # ── Encode ────────────────────────────────
        try:
            result = base64_encode(
                str(text),
                encoding=encoding,
                chunk_size=chunk_size,
            )
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(
            {
                "encoded": result["encoded"],
                "encoding": result["encoding"],
                "original_size_kb": result["original_size_kb"],
                "encoded_size_kb": result["encoded_size_kb"],
                "overhead_percent": result["overhead_percent"],
            }
        )


class Base64EncodeFileView(APIView):
    """
    POST /api/tools/base64/encode/file/
    Encode file → Base64.

    Form fields:
        file       : any file                        (required)
        encoding   : standard | url_safe | mime      (default: standard)
        chunk_size : split output every N chars      (default: 0)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        encoding = request.data.get("encoding", "standard")
        chunk_size = request.data.get("chunk_size", "0")

        # ── Validate ──────────────────────────────
        if not file:
            return Response(
                {"error": "No file provided."},
                status=400,
            )
        if encoding not in ("standard", "url_safe", "mime"):
            return Response(
                {"error": "encoding must be: standard, url_safe, or mime."},
                status=400,
            )
        try:
            chunk_size = int(chunk_size)
        except ValueError:
            return Response(
                {"error": "chunk_size must be an integer."},
                status=400,
            )

        # ── Encode ────────────────────────────────
        try:
            result = base64_encode(
                file,
                encoding=encoding,
                chunk_size=chunk_size,
                filename=file.name,
            )
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(
            {
                "encoded": result["encoded"],
                "encoding": result["encoding"],
                "filename": result["filename"],
                "mime_type": result["mime_type"],
                "data_uri": result["data_uri"],
                "original_size_kb": result["original_size_kb"],
                "encoded_size_kb": result["encoded_size_kb"],
                "overhead_percent": result["overhead_percent"],
            }
        )


class Base64DecodeView(APIView):
    """
    POST /api/tools/base64/decode/
    Decode Base64 → text or file download.

    JSON body:
        {
            "text"         : "SGVsbG8sIFdvcmxkIQ==",
            "encoding"     : "standard",
            "as_text"      : true,
            "text_encoding": "utf-8",
            "output"       : "json | file"
        }
    """

    parser_classes = [JSONParser, MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        text = request.data.get("text")
        encoding = request.data.get("encoding", "standard")
        as_text = request.data.get("as_text", True)
        text_encoding = request.data.get("text_encoding", "utf-8")
        output = request.data.get("output", "json")
        filename = request.data.get("filename", "decoded_file")

        # ── Validate ──────────────────────────────
        if not text:
            return Response(
                {"error": '"text" field is required.'},
                status=400,
            )
        if encoding not in ("standard", "url_safe"):
            return Response(
                {"error": "encoding must be: standard or url_safe."},
                status=400,
            )
        if output not in ("json", "file"):
            return Response(
                {"error": "output must be: json or file."},
                status=400,
            )
        if isinstance(as_text, str):
            as_text = as_text.lower() == "true"

        # ── Decode ────────────────────────────────
        try:
            result = base64_decode(
                text,
                encoding=encoding,
                as_text=as_text,
                text_encoding=text_encoding,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return file download ───────────────────
        if output == "file":
            response = HttpResponse(
                result["decoded_bytes"],
                content_type="application/octet-stream",
            )
            response["Content-Disposition"] = f'attachment; filename="{filename}"'
            response["Content-Length"] = result["decoded_size"]
            return response

        # ── Return JSON ───────────────────────────
        return Response(
            {
                "decoded_text": result["decoded_text"],
                "is_text": result["is_text"],
                "encoding": result["encoding"],
                "original_size_kb": result["original_size_kb"],
                "decoded_size_kb": result["decoded_size_kb"],
                "text_encoding": result["text_encoding"],
            }
        )


class Base64ValidateView(APIView):
    """
    POST /api/tools/base64/validate/
    Validate whether a string is valid Base64.

    JSON body:
        {
            "text": "SGVsbG8sIFdvcmxkIQ=="
        }
    """

    parser_classes = [JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        text = request.data.get("text")

        if not text:
            return Response(
                {"error": '"text" field is required.'},
                status=400,
            )

        try:
            result = base64_validate(str(text))
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(result)


class UUIDGeneratorView(APIView):
    """
    POST /api/tools/uuid/generate/
    Generate UUID(s) with full options.

    JSON body:
        {
            "version"  : 4,
            "count"    : 1,
            "uppercase": false,
            "hyphens"  : true,
            "braces"   : false,
            "prefix"   : "",
            "suffix"   : "",
            "namespace": "dns",
            "name"     : "example.com",
            "seed"     : null
        }
    """

    parser_classes = [JSONParser, MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        version = request.data.get("version", 4)
        count = request.data.get("count", 1)
        uppercase = request.data.get("uppercase", False)
        hyphens = request.data.get("hyphens", True)
        braces = request.data.get("braces", False)
        prefix = request.data.get("prefix", "")
        suffix = request.data.get("suffix", "")
        namespace = request.data.get("namespace", None)
        name = request.data.get("name", None)
        seed = request.data.get("seed", None)

        # ── Parse + Validate ──────────────────────
        try:
            version = int(version)
            count = int(count)
        except (ValueError, TypeError):
            return Response(
                {"error": "version and count must be integers."},
                status=400,
            )

        if isinstance(uppercase, str):
            uppercase = uppercase.lower() == "true"
        if isinstance(hyphens, str):
            hyphens = hyphens.lower() == "true"
        if isinstance(braces, str):
            braces = braces.lower() == "true"
        if seed is not None:
            try:
                seed = int(seed)
            except (ValueError, TypeError):
                return Response(
                    {"error": "seed must be an integer."},
                    status=400,
                )

        # ── Generate ──────────────────────────────
        try:
            result = generate_uuid(
                version=version,
                count=count,
                uppercase=uppercase,
                hyphens=hyphens,
                braces=braces,
                prefix=str(prefix),
                suffix=str(suffix),
                namespace=namespace,
                name=name,
                seed=seed,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(result)


class UUIDValidateView(APIView):
    """
    POST /api/tools/uuid/validate/
    Validate a UUID and return its details.

    JSON body:
        {
            "uuid": "550e8400-e29b-41d4-a716-446655440000"
        }
    """

    parser_classes = [JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        value = request.data.get("uuid")

        if not value:
            return Response(
                {"error": '"uuid" field is required.'},
                status=400,
            )

        try:
            result = validate_uuid(str(value))
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(result)


class UUIDBulkView(APIView):
    """
    POST /api/tools/uuid/bulk/
    Generate bulk UUIDs in multiple export formats.

    JSON body:
        {
            "version": 4,
            "count"  : 10,
            "format" : "standard | csv | json | sql | array"
        }
    """

    parser_classes = [JSONParser, MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        version = request.data.get("version", 4)
        count = request.data.get("count", 10)
        format = request.data.get("format", "standard")
        output = request.data.get("output", "json")

        try:
            version = int(version)
            count = int(count)
        except (ValueError, TypeError):
            return Response(
                {"error": "version and count must be integers."},
                status=400,
            )

        if format not in ("standard", "csv", "json", "sql", "array"):
            return Response(
                {"error": "format must be: standard, csv, json, sql, or array."},
                status=400,
            )

        try:
            result = bulk_generate_uuids(
                version=version,
                count=count,
                format=format,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return file download ───────────────────
        if output == "file":
            ext_map = {
                "csv": ("text/csv", ".csv"),
                "json": ("application/json", ".json"),
                "sql": ("text/plain", ".sql"),
                "standard": ("text/plain", ".txt"),
                "array": ("application/json", ".json"),
            }
            content_type, ext = ext_map.get(format, ("text/plain", ".txt"))
            response = HttpResponse(
                result["export"],
                content_type=content_type,
            )
            response["Content-Disposition"] = (
                f'attachment; filename="uuids_{count}{ext}"'
            )
            return response

        return Response(result)


class FileEncryptView(APIView):
    """
    POST /api/tools/file/encrypt/
    Upload any file → returns encrypted .enc file.

    Form fields:
        file      : any file                             (required)
        password  : encryption password                 (required)
        algorithm : AES-256-GCM | AES-256-CBC | ChaCha20 (default: AES-256-GCM)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        password = request.data.get("password", "").strip()
        algorithm = request.data.get("algorithm", "AES-256-GCM")

        # ── Validate ──────────────────────────────
        if not file:
            return Response(
                {"error": "No file provided."},
                status=400,
            )
        if not password:
            return Response(
                {"error": "password is required."},
                status=400,
            )
        if len(password) < 6:
            return Response(
                {"error": "password must be at least 6 characters."},
                status=400,
            )
        if algorithm not in ("AES-256-GCM", "AES-256-CBC", "ChaCha20"):
            return Response(
                {"error": "algorithm must be: AES-256-GCM, AES-256-CBC, or ChaCha20."},
                status=400,
            )

        # ── Encrypt ───────────────────────────────
        try:
            result = encrypt_file(
                file,
                password=password,
                algorithm=algorithm,
                filename=file.name,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return encrypted file ──────────────────
        response = HttpResponse(
            result["encrypted_bytes"],
            content_type="application/octet-stream",
        )
        response["Content-Disposition"] = (
            f'attachment; filename="{result["output_filename"]}"'
        )
        response["Content-Length"] = result["encrypted_size"]
        response["X-Algorithm"] = result["algorithm"]
        response["X-Original-Size-KB"] = result["original_size_kb"]
        response["X-Encrypted-Size-KB"] = result["encrypted_size_kb"]
        return response


class FileDecryptView(APIView):
    """
    POST /api/tools/file/decrypt/
    Upload encrypted .enc file + password → returns original file.

    Form fields:
        file     : encrypted .enc file    (required)
        password : decryption password    (required)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        password = request.data.get("password", "").strip()

        # ── Validate ──────────────────────────────
        if not file:
            return Response(
                {"error": "No file provided."},
                status=400,
            )
        if not password:
            return Response(
                {"error": "password is required."},
                status=400,
            )

        # ── Decrypt ───────────────────────────────
        try:
            result = decrypt_file(file, password=password)
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return original file ───────────────────
        response = HttpResponse(
            result["decrypted_bytes"],
            content_type="application/octet-stream",
        )
        response["Content-Disposition"] = (
            f'attachment; filename="{result["original_filename"]}"'
        )
        response["Content-Length"] = result["original_size"]
        response["X-Algorithm"] = result["algorithm"]
        response["X-Original-Size-KB"] = result["original_size_kb"]
        return response


class FileHashView(APIView):
    """
    POST /api/tools/file/hash/
    Upload a file or send text → returns hash(es).

    Form fields (file):
        file       : any file                           (required)
        algorithms : comma-separated hash algorithms   (default: all)
                     md5,sha1,sha256,sha512,sha3_256

    JSON body (text):
        {
            "text"      : "Hello, World!",
            "algorithms": ["sha256", "md5"]
        }
    """

    parser_classes = [MultiPartParser, JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file", None)
        text = request.data.get("text", None)
        algorithms = request.data.get("algorithms", None)

        # ── Validate ──────────────────────────────
        if not file and not text:
            return Response(
                {"error": 'Provide either "file" or "text".'},
                status=400,
            )

        # ── Parse algorithms ──────────────────────
        parsed_algos = None
        if algorithms:
            if isinstance(algorithms, str):
                parsed_algos = [
                    a.strip().lower() for a in algorithms.split(",") if a.strip()
                ]
            elif isinstance(algorithms, list):
                parsed_algos = [a.lower() for a in algorithms]

        # ── Hash ──────────────────────────────────
        try:
            source = file if file else text
            result = get_file_hash(source, algorithms_list=parsed_algos)
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(
            {
                "filename": file.name if file else None,
                "size_kb": result["size_kb"],
                "size_mb": result["size_mb"],
                "hashes": result["hashes"],
            }
        )


class PasswordGeneratorView(APIView):
    """
    POST /api/tools/password-generate/
    Generate secure password(s) with full options.

    JSON body:
        {
            "length"          : 16,
            "count"           : 1,
            "uppercase"       : true,
            "lowercase"       : true,
            "digits"          : true,
            "symbols"         : true,
            "exclude_chars"   : "",
            "exclude_similar" : false,
            "exclude_ambiguous": false,
            "custom_chars"    : "",
            "prefix"          : "",
            "suffix"          : "",
            "no_repeat"       : false
        }
    """

    parser_classes = [JSONParser, MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        length = request.data.get("length", 16)
        count = request.data.get("count", 1)
        uppercase = request.data.get("uppercase", True)
        lowercase = request.data.get("lowercase", True)
        digits = request.data.get("digits", True)
        symbols = request.data.get("symbols", True)
        exclude_chars = request.data.get("exclude_chars", "")
        exclude_similar = request.data.get("exclude_similar", False)
        exclude_ambiguous = request.data.get("exclude_ambiguous", False)
        custom_chars = request.data.get("custom_chars", "")
        prefix = request.data.get("prefix", "")
        suffix = request.data.get("suffix", "")
        no_repeat = request.data.get("no_repeat", False)

        # ── Parse types ───────────────────────────
        try:
            length = int(length)
            count = int(count)
        except (ValueError, TypeError):
            return Response(
                {"error": "length and count must be integers."},
                status=400,
            )

        def to_bool(val):
            if isinstance(val, bool):
                return val
            return str(val).lower() == "true"

        uppercase = to_bool(uppercase)
        lowercase = to_bool(lowercase)
        digits = to_bool(digits)
        symbols = to_bool(symbols)
        exclude_similar = to_bool(exclude_similar)
        exclude_ambiguous = to_bool(exclude_ambiguous)
        no_repeat = to_bool(no_repeat)

        # ── Generate ──────────────────────────────
        try:
            result = generate_password(
                length=length,
                count=count,
                uppercase=uppercase,
                lowercase=lowercase,
                digits=digits,
                symbols=symbols,
                exclude_chars=str(exclude_chars),
                exclude_similar=exclude_similar,
                exclude_ambiguous=exclude_ambiguous,
                custom_chars=str(custom_chars),
                prefix=str(prefix),
                suffix=str(suffix),
                no_repeat=no_repeat,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(result)


class PasswordStrengthView(APIView):
    """
    POST /api/tools/password-strength/
    Check strength of an existing password.

    JSON body:
        {
            "password": "MyPassword123!"
        }
    """

    parser_classes = [JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        password = request.data.get("password")

        if not password:
            return Response(
                {"error": '"password" field is required.'},
                status=400,
            )

        try:
            result = check_password_strength(str(password))
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(result)


class PassphraseGeneratorView(APIView):
    """
    POST /api/tools/passphrase-generate/
    Generate memorable passphrase(s).

    JSON body:
        {
            "words"        : 4,
            "count"        : 1,
            "separator"    : "-",
            "capitalize"   : true,
            "include_digit": true
        }
    """

    parser_classes = [JSONParser, MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        words = request.data.get("words", 4)
        count = request.data.get("count", 1)
        separator = request.data.get("separator", "-")
        capitalize = request.data.get("capitalize", True)
        include_digit = request.data.get("include_digit", True)

        # ── Parse ─────────────────────────────────
        try:
            words = int(words)
            count = int(count)
        except (ValueError, TypeError):
            return Response(
                {"error": "words and count must be integers."},
                status=400,
            )

        def to_bool(val):
            if isinstance(val, bool):
                return val
            return str(val).lower() == "true"

        capitalize = to_bool(capitalize)
        include_digit = to_bool(include_digit)

        # ── Generate ──────────────────────────────
        try:
            result = generate_passphrase(
                words=words,
                count=count,
                separator=str(separator),
                capitalize=capitalize,
                include_digit=include_digit,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(result)


class HashGeneratorView(APIView):
    """
    POST /api/tools/hash-generate/
    Generate hash(es) from text or file.

    JSON body (text):
        {
            "text"        : "Hello, World!",
            "algorithms"  : ["sha256", "md5"],
            "encoding"    : "utf-8",
            "hmac_key"    : null,
            "output_format": "hex"
        }

    Form fields (file):
        file          : any file                          (required)
        algorithms    : comma-separated algorithms        (default: all)
        hmac_key      : HMAC secret key                  (optional)
        output_format : hex | base64 | base64url | int   (default: hex)
    """

    parser_classes = [MultiPartParser, JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file", None)
        text = request.data.get("text", None)
        algorithms = request.data.get("algorithms", None)
        encoding = request.data.get("encoding", "utf-8")
        hmac_key = request.data.get("hmac_key", None)
        output_format = request.data.get("output_format", "hex")

        # ── Validate ──────────────────────────────
        if not file and text is None:
            return Response(
                {"error": 'Provide either "file" or "text".'},
                status=400,
            )

        if output_format not in ("hex", "base64", "base64url", "int"):
            return Response(
                {"error": "output_format must be: hex, base64, base64url, or int."},
                status=400,
            )

        # ── Parse algorithms ──────────────────────
        parsed_algos = None
        if algorithms:
            if isinstance(algorithms, str):
                parsed_algos = [
                    a.strip().lower() for a in algorithms.split(",") if a.strip()
                ]
            elif isinstance(algorithms, list):
                parsed_algos = [a.lower() for a in algorithms]

        # ── Generate hash ──────────────────────────
        try:
            source = file if file else str(text)
            result = generate_hash(
                source,
                algorithms_list=parsed_algos,
                encoding=encoding,
                hmac_key=hmac_key,
                output_format=output_format,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(
            {
                "filename": file.name if file else None,
                "input_type": result["input_type"],
                "input_size_kb": result["input_size_kb"],
                "encoding": result["encoding"],
                "output_format": result["output_format"],
                "is_hmac": result["is_hmac"],
                "algorithm_count": result["algorithm_count"],
                "hashes": result["hashes"],
            }
        )


class HashCompareView(APIView):
    """
    POST /api/tools/hash-compare/
    Securely compare two hash strings.

    JSON body:
        {
            "hash1": "a665a45920422f9d...",
            "hash2": "a665a45920422f9d..."
        }
    """

    parser_classes = [JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        hash1 = request.data.get("hash1")
        hash2 = request.data.get("hash2")

        if not hash1 or not hash2:
            return Response(
                {"error": 'Both "hash1" and "hash2" are required.'},
                status=400,
            )

        try:
            result = compare_hashes(str(hash1), str(hash2))
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(result)


class FileChecksumView(APIView):
    """
    POST /api/tools/file-checksum/
    Generate a file checksum for integrity verification.

    Form fields:
        file      : any file                    (required)
        algorithm : hash algorithm              (default: sha256)
                    md5|sha1|sha256|sha512|...
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        algorithm = request.data.get("algorithm", "sha256")

        if not file:
            return Response(
                {"error": "No file provided."},
                status=400,
            )

        try:
            result = generate_checksum(file, algorithm=algorithm)
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(
            {
                "filename": file.name,
                "checksum": result["checksum"],
                "algorithm": result["algorithm"],
                "size_kb": result["size_kb"],
                "size": result["size"],
            }
        )


class WordToJPGView(APIView):
    """
    POST /api/convert/word-to-jpg/
    Upload a Word file (.docx/.doc) → returns JPG image(s).

    Form fields:
        file     : Word file (.docx / .doc)              (required)
        dpi      : image resolution 72-600               (default: 200)
        quality  : JPG quality 1-95                      (default: 85)
        pages    : comma-separated page numbers          (default: all)
                   e.g. "1" or "1,3,5"
        output   : single | zip | base64                 (default: zip)
                   - single  → first (or specific) page as .jpg
                   - zip     → all pages as .zip of .jpg files
                   - base64  → JSON with base64 encoded images
        filename : output filename                       (optional)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        dpi = request.data.get("dpi", "200")
        quality = request.data.get("quality", "85")
        pages = request.data.get("pages", None)
        output = request.data.get("output", "zip")
        filename = request.data.get("filename", None)

        # ── Validate ──────────────────────────────
        if not file:
            return Response(
                {"error": "No file provided."},
                status=400,
            )

        if not file.name.lower().endswith((".docx", ".doc")):
            return Response(
                {"error": "Only .docx or .doc files are accepted."},
                status=400,
            )

        if output not in ("single", "zip", "base64"):
            return Response(
                {"error": "output must be: single, zip, or base64."},
                status=400,
            )

        # ── Parse DPI ─────────────────────────────
        try:
            dpi = int(dpi)
            if not (72 <= dpi <= 600):
                raise ValueError
        except ValueError:
            return Response(
                {"error": "dpi must be an integer between 72 and 600."},
                status=400,
            )

        # ── Parse quality ─────────────────────────
        try:
            quality = int(quality)
            if not (1 <= quality <= 95):
                raise ValueError
        except ValueError:
            return Response(
                {"error": "quality must be an integer between 1 and 95."},
                status=400,
            )

        # ── Parse pages ───────────────────────────
        parsed_pages = None
        if pages:
            try:
                parsed_pages = [int(p.strip()) for p in pages.split(",") if p.strip()]
            except ValueError:
                return Response(
                    {"error": 'pages must be comma-separated integers: e.g. "1,2,3"'},
                    status=400,
                )

        # ── Base filename ─────────────────────────
        base_name = (
            file.name.replace(".docx", "")
            .replace(".doc", "")
            .replace(".DOCX", "")
            .replace(".DOC", "")
        )

        # ── Convert ───────────────────────────────
        try:
            jpg_pages = word_to_jpg(
                file,
                filename=file.name,
                dpi=dpi,
                quality=quality,
                pages=parsed_pages,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except RuntimeError as e:
            return Response({"error": str(e)}, status=500)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        if not jpg_pages:
            return Response(
                {"error": "No pages found in document."},
                status=400,
            )

        # ── Output: single JPG ─────────────────────
        if output == "single":
            first = jpg_pages[0]
            out_name = filename or f'{base_name}_page_{first["page"]}.jpg'
            if not out_name.endswith(".jpg"):
                out_name += ".jpg"

            response = HttpResponse(first["bytes"], content_type="image/jpeg")
            response["Content-Disposition"] = f'attachment; filename="{out_name}"'
            response["Content-Length"] = len(first["bytes"])
            response["X-Page"] = first["page"]
            response["X-Width"] = first["width"]
            response["X-Height"] = first["height"]
            response["X-Size-KB"] = first["size_kb"]
            return response

        # ── Output: ZIP of all pages ───────────────
        if output == "zip":
            import zipfile

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for p in jpg_pages:
                    zf.writestr(
                        f'{base_name}_page_{p["page"]}.jpg',
                        p["bytes"],
                    )
            zip_buffer.seek(0)

            out_name = filename or f"{base_name}_pages.zip"
            if not out_name.endswith(".zip"):
                out_name += ".zip"

            response = HttpResponse(
                zip_buffer.read(),
                content_type="application/zip",
            )
            response["Content-Disposition"] = f'attachment; filename="{out_name}"'
            response["X-Total-Pages"] = len(jpg_pages)
            return response

        # ── Output: base64 JSON ────────────────────
        if output == "base64":
            data = [
                {
                    "page": p["page"],
                    "filename": f'{base_name}_page_{p["page"]}.jpg',
                    "width": p["width"],
                    "height": p["height"],
                    "size_kb": p["size_kb"],
                    "base64": base64.b64encode(p["bytes"]).decode("utf-8"),
                }
                for p in jpg_pages
            ]

            return Response(
                {
                    "total_pages": len(data),
                    "dpi": dpi,
                    "quality": quality,
                    "format": "JPG",
                    "pages": data,
                }
            )


class WordToTXTView(APIView):
    """
    POST /api/convert/word-to-txt/
    Upload a Word file (.docx/.doc) → returns plain text.

    Form fields:
        file             : Word file (.docx / .doc)    (required)
        include_headers  : true | false                (default: true)
        include_tables   : true | false                (default: true)
        include_comments : true | false                (default: false)
        preserve_spacing : true | false                (default: true)
        page_separator   : separator string            (default: empty)
        encoding         : utf-8 | ascii | latin-1     (default: utf-8)
        output           : json | file                 (default: json)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        include_headers = request.data.get("include_headers", "true")
        include_tables = request.data.get("include_tables", "true")
        include_comments = request.data.get("include_comments", "false")
        preserve_spacing = request.data.get("preserve_spacing", "true")
        page_separator = request.data.get("page_separator", "")
        encoding = request.data.get("encoding", "utf-8")
        output = request.data.get("output", "json")

        # ── Validate ──────────────────────────────
        if not file:
            return Response(
                {"error": "No file provided."},
                status=400,
            )

        if not file.name.lower().endswith((".docx", ".doc")):
            return Response(
                {"error": "Only .docx or .doc files are accepted."},
                status=400,
            )

        if output not in ("json", "file"):
            return Response(
                {"error": "output must be: json or file."},
                status=400,
            )

        if encoding not in ("utf-8", "ascii", "latin-1", "utf-16"):
            return Response(
                {"error": "encoding must be: utf-8, ascii, latin-1, or utf-16."},
                status=400,
            )

        # ── Parse booleans ────────────────────────
        def to_bool(val):
            if isinstance(val, bool):
                return val
            return str(val).lower() == "true"

        include_headers = to_bool(include_headers)
        include_tables = to_bool(include_tables)
        include_comments = to_bool(include_comments)
        preserve_spacing = to_bool(preserve_spacing)

        # ── Convert ───────────────────────────────
        try:
            result = word_to_txt(
                file,
                filename=file.name,
                include_headers=include_headers,
                include_tables=include_tables,
                include_comments=include_comments,
                preserve_spacing=preserve_spacing,
                page_separator=page_separator,
                encoding=encoding,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Output: file download ──────────────────
        if output == "file":
            txt_filename = (
                file.name.replace(".docx", ".txt")
                .replace(".doc", ".txt")
                .replace(".DOCX", ".txt")
                .replace(".DOC", ".txt")
            )

            response = HttpResponse(
                result["text"].encode(encoding),
                content_type=f"text/plain; charset={encoding}",
            )
            response["Content-Disposition"] = f'attachment; filename="{txt_filename}"'
            response["Content-Length"] = result["size_text"]
            response["X-Word-Count"] = result["words"]
            response["X-Char-Count"] = result["chars"]
            response["X-Paragraph-Count"] = result["paragraphs"]
            response["X-Table-Count"] = result["tables"]
            response["X-Heading-Count"] = result["headings"]
            return response

        # ── Output: JSON response ──────────────────
        return Response(
            {
                "text": result["text"],
                "paragraphs": result["paragraphs"],
                "words": result["words"],
                "chars": result["chars"],
                "tables": result["tables"],
                "headings": result["headings"],
                "size_original_kb": result["size_original_kb"],
                "size_text_kb": result["size_text_kb"],
                "encoding": result["encoding"],
            }
        )


class TXTToWordView(APIView):
    """
    POST /api/convert/txt-to-word/
    Upload .txt file or send raw text → returns .docx file.

    Form fields (file upload):
        file            : .txt file                          (required)
        title           : document title                    (optional)
        font_name       : Calibri|Arial|Times New Roman     (default: Calibri)
        font_size       : font size in pt                   (default: 11)
        line_spacing    : 1.0 | 1.15 | 1.5 | 2.0          (default: 1.15)
        detect_headings : true | false                      (default: true)
        detect_lists    : true | false                      (default: true)
        detect_tables   : true | false                      (default: true)
        page_size       : A4 | Letter | Legal | A3          (default: A4)
        encoding        : utf-8 | ascii | latin-1           (default: utf-8)
        output_filename : custom output filename            (optional)

    JSON body (raw text):
        {
            "text"           : "Your plain text here...",
            "title"          : "My Document",
            "font_name"      : "Calibri",
            "font_size"      : 11,
            "line_spacing"   : 1.15,
            "detect_headings": true,
            "detect_lists"   : true,
            "detect_tables"  : true,
            "page_size"      : "A4"
        }
    """

    parser_classes = [MultiPartParser, JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file", None)
        text = request.data.get("text", None)
        title = request.data.get("title", "")
        font_name = request.data.get("font_name", "Calibri")
        font_size = request.data.get("font_size", 11)
        line_spacing = request.data.get("line_spacing", 1.15)
        detect_headings = request.data.get("detect_headings", True)
        detect_lists = request.data.get("detect_lists", True)
        detect_tables = request.data.get("detect_tables", True)
        page_size = request.data.get("page_size", "A4")
        encoding = request.data.get("encoding", "utf-8")
        output_filename = request.data.get("output_filename", None)

        # ── Validate ──────────────────────────────
        if not file and text is None:
            return Response(
                {"error": 'Provide either "file" (.txt) or "text" (raw string).'},
                status=400,
            )

        if file and not file.name.lower().endswith((".txt", ".text", ".md")):
            return Response(
                {"error": "Only .txt, .text, or .md files are accepted."},
                status=400,
            )

        # ── Parse types ───────────────────────────
        try:
            font_size = int(font_size)
            line_spacing = float(line_spacing)
        except (ValueError, TypeError):
            return Response(
                {"error": "font_size must be integer, line_spacing must be float."},
                status=400,
            )

        if not (6 <= font_size <= 72):
            return Response(
                {"error": "font_size must be between 6 and 72."},
                status=400,
            )

        if not (1.0 <= line_spacing <= 3.0):
            return Response(
                {"error": "line_spacing must be between 1.0 and 3.0."},
                status=400,
            )

        valid_fonts = (
            "Calibri",
            "Arial",
            "Times New Roman",
            "Georgia",
            "Verdana",
            "Helvetica",
            "Courier New",
            "Tahoma",
            "Trebuchet MS",
        )
        if font_name not in valid_fonts:
            return Response(
                {"error": f'font_name must be one of: {", ".join(valid_fonts)}'},
                status=400,
            )

        if page_size not in ("A4", "Letter", "Legal", "A3"):
            return Response(
                {"error": "page_size must be: A4, Letter, Legal, or A3."},
                status=400,
            )

        # ── Parse booleans ─────────────────────────
        def to_bool(val):
            if isinstance(val, bool):
                return val
            return str(val).lower() == "true"

        detect_headings = to_bool(detect_headings)
        detect_lists = to_bool(detect_lists)
        detect_tables = to_bool(detect_tables)

        # ── Source + filename ──────────────────────
        source = file if file else str(text)
        src_name = file.name if file else "document.txt"

        if not output_filename:
            output_filename = (
                src_name.replace(".txt", ".docx")
                .replace(".text", ".docx")
                .replace(".md", ".docx")
            )
        if not output_filename.endswith(".docx"):
            output_filename += ".docx"

        # ── Convert ───────────────────────────────
        try:
            docx_bytes = txt_to_word(
                source,
                filename=src_name,
                title=str(title),
                font_name=font_name,
                font_size=font_size,
                line_spacing=line_spacing,
                detect_headings=detect_headings,
                detect_lists=detect_lists,
                detect_tables=detect_tables,
                page_size=page_size,
                encoding=encoding,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return .docx download ──────────────────
        response = HttpResponse(
            docx_bytes,
            content_type=(
                "application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.document"
            ),
        )
        response["Content-Disposition"] = f'attachment; filename="{output_filename}"'
        response["Content-Length"] = len(docx_bytes)
        return response


class MergeDOCXView(APIView):
    """
    POST /api/tools/merge-docx/
    Upload 2-20 Word files → returns merged .docx file.

    Form fields:
        files[]         : multiple .docx files              (required, min 2)
        page_break      : true | false                      (default: true)
        add_toc         : true | false                      (default: false)
        preserve_styles : true | false                      (default: true)
        separator_text  : text between docs                 (optional)
                          use {n} for doc number
                          use {filename} for filename
        output_filename : custom output filename            (optional)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        files = request.FILES.getlist("files[]")
        page_break = request.data.get("page_break", "true")
        add_toc = request.data.get("add_toc", "false")
        preserve_styles = request.data.get("preserve_styles", "true")
        separator_text = request.data.get("separator_text", "")
        output_filename = request.data.get("output_filename", "merged.docx")

        # ── Also support files without [] ──────────
        if not files:
            files = request.FILES.getlist("files")

        # ── Validate ──────────────────────────────
        if not files:
            return Response(
                {"error": "No files provided. Use files[] field with multiple files."},
                status=400,
            )

        if len(files) < 2:
            return Response(
                {"error": "At least 2 files are required to merge."},
                status=400,
            )

        if len(files) > 20:
            return Response(
                {"error": "Maximum 20 files can be merged at once."},
                status=400,
            )

        # ── Validate file types ────────────────────
        for f in files:
            if not f.name.lower().endswith((".docx", ".doc")):
                return Response(
                    {"error": f'"{f.name}" is not a .docx or .doc file.'},
                    status=400,
                )

        # ── Parse booleans ─────────────────────────
        def to_bool(val):
            if isinstance(val, bool):
                return val
            return str(val).lower() == "true"

        page_break = to_bool(page_break)
        add_toc = to_bool(add_toc)
        preserve_styles = to_bool(preserve_styles)

        # ── Output filename ───────────────────────
        if not output_filename.endswith(".docx"):
            output_filename += ".docx"

        # ── Merge ─────────────────────────────────
        try:
            merged_bytes = merge_docx(
                sources=files,
                page_break=page_break,
                add_toc=add_toc,
                preserve_styles=preserve_styles,
                separator_text=separator_text,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Return merged file ─────────────────────
        response = HttpResponse(
            merged_bytes,
            content_type=(
                "application/vnd.openxmlformats-officedocument"
                ".wordprocessingml.document"
            ),
        )
        response["Content-Disposition"] = f'attachment; filename="{output_filename}"'
        response["Content-Length"] = len(merged_bytes)
        response["X-Files-Merged"] = len(files)
        response["X-File-Names"] = ", ".join(f.name for f in files)
        return response


class SplitDOCXView(APIView):
    """
    POST /api/tools/split-docx/
    Upload a Word file → returns split .docx files as .zip.

    Form fields:
        file          : Word file (.docx / .doc)           (required)
        split_by      : page | heading | chunk | range     (default: page)
        heading_level : 1-6 (for heading mode)             (default: 1)
        chunk_size    : paragraphs per chunk               (default: 10)
        pages         : JSON array for range mode          (optional)
                        e.g. [[1,3],[4,6]] or [1,3,5]
        output        : zip | json                         (default: zip)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        split_by = request.data.get("split_by", "page")
        heading_level = request.data.get("heading_level", "1")
        chunk_size = request.data.get("chunk_size", "10")
        pages_raw = request.data.get("pages", None)
        output = request.data.get("output", "zip")

        # ── Validate ──────────────────────────────
        if not file:
            return Response(
                {"error": "No file provided."},
                status=400,
            )

        if not file.name.lower().endswith((".docx", ".doc")):
            return Response(
                {"error": "Only .docx or .doc files are accepted."},
                status=400,
            )

        if split_by not in ("page", "heading", "chunk", "range"):
            return Response(
                {"error": "split_by must be: page, heading, chunk, or range."},
                status=400,
            )

        if output not in ("zip", "json"):
            return Response(
                {"error": "output must be: zip or json."},
                status=400,
            )

        # ── Parse heading_level ───────────────────
        try:
            heading_level = int(heading_level)
            if not (1 <= heading_level <= 6):
                raise ValueError
        except (ValueError, TypeError):
            return Response(
                {"error": "heading_level must be an integer between 1 and 6."},
                status=400,
            )

        # ── Parse chunk_size ──────────────────────
        try:
            chunk_size = int(chunk_size)
            if chunk_size < 1:
                raise ValueError
        except (ValueError, TypeError):
            return Response(
                {"error": "chunk_size must be a positive integer."},
                status=400,
            )

        # ── Parse pages (range mode) ───────────────
        parsed_pages = None
        if pages_raw:
            import json as _json

            try:
                if isinstance(pages_raw, str):
                    parsed_pages = _json.loads(pages_raw)
                elif isinstance(pages_raw, list):
                    parsed_pages = pages_raw
            except Exception:
                return Response(
                    {"error": "pages must be valid JSON: [[1,3],[4,6]] or [1,3,5]"},
                    status=400,
                )

        if split_by == "range" and not parsed_pages:
            return Response(
                {"error": "pages is required for range mode. e.g. [[1,3],[4,6]]"},
                status=400,
            )

        # ── Split ─────────────────────────────────
        try:
            parts = split_docx(
                file,
                split_by=split_by,
                pages=parsed_pages,
                heading_level=heading_level,
                chunk_size=chunk_size,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        base_name = (
            file.name.replace(".docx", "")
            .replace(".doc", "")
            .replace(".DOCX", "")
            .replace(".DOC", "")
        )

        # ── Output: JSON (metadata only) ──────────
        if output == "json":
            return Response(
                {
                    "total_parts": len(parts),
                    "split_by": split_by,
                    "parts": [
                        {
                            "index": p["index"],
                            "filename": p["filename"],
                            "title": p["title"],
                            "paragraphs": p["paragraphs"],
                            "size_kb": p["size_kb"],
                        }
                        for p in parts
                    ],
                }
            )

        # ── Output: ZIP of all parts (default) ────
        import zipfile

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for part in parts:
                zf.writestr(part["filename"], part["bytes"])

        zip_buffer.seek(0)

        response = HttpResponse(
            zip_buffer.read(),
            content_type="application/zip",
        )
        response["Content-Disposition"] = (
            f'attachment; filename="{base_name}_split.zip"'
        )
        response["X-Total-Parts"] = len(parts)
        response["X-Split-By"] = split_by
        return response


class TextCompareView(APIView):
    """
    POST /api/tools/text-compare
    Compare two raw text strings.

    JSON body:
        {
            "text1"             : "original text...",
            "text2"             : "modified text...",
            "compare_mode"      : "line | word | char | sentence",
            "ignore_case"       : false,
            "ignore_whitespace" : false,
            "ignore_blank_lines": false,
            "context_lines"     : 3,
            "output_format"     : "unified | html | json | side_by_side"
        }
    """

    parser_classes = [JSONParser]
    permission_classes = [AllowAny]

    def post(self, request):
        text1 = request.data.get("text1")
        text2 = request.data.get("text2")
        compare_mode = request.data.get("compare_mode", "line")
        ignore_case = request.data.get("ignore_case", False)
        ignore_whitespace = request.data.get("ignore_whitespace", False)
        ignore_blank_lines = request.data.get("ignore_blank_lines", False)
        context_lines = request.data.get("context_lines", 3)
        output_format = request.data.get("output_format", "unified")

        # ── Validate ──────────────────────────────
        if text1 is None or text2 is None:
            return Response(
                {"error": 'Both "text1" and "text2" are required.'},
                status=400,
            )

        if compare_mode not in ("line", "word", "char", "sentence"):
            return Response(
                {"error": "compare_mode must be: line, word, char, or sentence."},
                status=400,
            )

        if output_format not in ("unified", "html", "json", "side_by_side"):
            return Response(
                {
                    "error": "output_format must be: unified, html, json, or side_by_side."
                },
                status=400,
            )

        # ── Parse types ───────────────────────────
        def to_bool(val):
            if isinstance(val, bool):
                return val
            return str(val).lower() == "true"

        ignore_case = to_bool(ignore_case)
        ignore_whitespace = to_bool(ignore_whitespace)
        ignore_blank_lines = to_bool(ignore_blank_lines)

        try:
            context_lines = int(context_lines)
            if not (0 <= context_lines <= 10):
                raise ValueError
        except (ValueError, TypeError):
            return Response(
                {"error": "context_lines must be an integer between 0 and 10."},
                status=400,
            )

        # ── Compare ───────────────────────────────
        try:
            result = compare_texts(
                str(text1),
                str(text2),
                compare_mode=compare_mode,
                ignore_case=ignore_case,
                ignore_whitespace=ignore_whitespace,
                ignore_blank_lines=ignore_blank_lines,
                context_lines=context_lines,
                output_format=output_format,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(result)


class FileCompareView(APIView):
    """
    POST /api/tools/file-compare/
    Upload two text files → returns comparison result.

    Form fields:
        file1              : first file  (original)    (required)
        file2              : second file (modified)    (required)
        compare_mode       : line | word | char | sentence  (default: line)
        ignore_case        : true | false              (default: false)
        ignore_whitespace  : true | false              (default: false)
        ignore_blank_lines : true | false              (default: false)
        context_lines      : 0-10                      (default: 3)
        output_format      : unified | html | json | side_by_side (default: unified)
        encoding           : utf-8 | ascii | latin-1   (default: utf-8)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file1 = request.FILES.get("file1")
        file2 = request.FILES.get("file2")
        compare_mode = request.data.get("compare_mode", "line")
        ignore_case = request.data.get("ignore_case", "false")
        ignore_whitespace = request.data.get("ignore_whitespace", "false")
        ignore_blank_lines = request.data.get("ignore_blank_lines", "false")
        context_lines = request.data.get("context_lines", "3")
        output_format = request.data.get("output_format", "unified")
        encoding = request.data.get("encoding", "utf-8")

        # ── Validate ──────────────────────────────
        if not file1 or not file2:
            return Response(
                {"error": 'Both "file1" and "file2" are required.'},
                status=400,
            )

        if compare_mode not in ("line", "word", "char", "sentence"):
            return Response(
                {"error": "compare_mode must be: line, word, char, or sentence."},
                status=400,
            )

        if output_format not in ("unified", "html", "json", "side_by_side"):
            return Response(
                {
                    "error": "output_format must be: unified, html, json, or side_by_side."
                },
                status=400,
            )

        # ── Parse types ───────────────────────────
        def to_bool(val):
            if isinstance(val, bool):
                return val
            return str(val).lower() == "true"

        ignore_case = to_bool(ignore_case)
        ignore_whitespace = to_bool(ignore_whitespace)
        ignore_blank_lines = to_bool(ignore_blank_lines)

        try:
            context_lines = int(context_lines)
        except (ValueError, TypeError):
            return Response(
                {"error": "context_lines must be an integer."},
                status=400,
            )

        # ── Compare ───────────────────────────────
        try:
            result = compare_files(
                file1,
                file2,
                encoding=encoding,
                compare_mode=compare_mode,
                ignore_case=ignore_case,
                ignore_whitespace=ignore_whitespace,
                output_format=output_format,
                context_lines=context_lines,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        return Response(result)


class MergePDFView(APIView):
    """
    POST /api/tools/pdf-merge/
    Upload 2-50 PDF files → returns merged PDF.

    Form fields:
        files[]          : multiple PDF files              (required, min 2)
        add_bookmarks    : true | false                    (default: true)
        add_page_numbers : true | false                    (default: false)
        password         : output PDF password             (optional)
        output_filename  : custom output filename          (optional)
        output           : file | json                     (default: file)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        files = request.FILES.getlist("files[]")
        add_bookmarks = request.data.get("add_bookmarks", "true")
        add_page_numbers = request.data.get("add_page_numbers", "false")
        password = request.data.get("password", None)
        output_filename = request.data.get("output_filename", "merged.pdf")
        output = request.data.get("output", "file")

        # ── Also support without [] ────────────────
        if not files:
            files = request.FILES.getlist("files")

        # ── Validate ──────────────────────────────
        if not files:
            return Response(
                {"error": "No files provided. Use files[] field."},
                status=400,
            )
        if len(files) < 2:
            return Response(
                {"error": "At least 2 PDF files are required."},
                status=400,
            )
        if len(files) > 50:
            return Response(
                {"error": "Maximum 50 files can be merged at once."},
                status=400,
            )

        # ── Validate file types ────────────────────
        for f in files:
            if not f.name.lower().endswith(".pdf"):
                return Response(
                    {"error": f'"{f.name}" is not a PDF file.'},
                    status=400,
                )

        if output not in ("file", "json"):
            return Response(
                {"error": "output must be: file or json."},
                status=400,
            )

        # ── Parse booleans ─────────────────────────
        def to_bool(val):
            if isinstance(val, bool):
                return val
            return str(val).lower() == "true"

        add_bookmarks = to_bool(add_bookmarks)
        add_page_numbers = to_bool(add_page_numbers)

        # ── Output filename ───────────────────────
        if not output_filename.endswith(".pdf"):
            output_filename += ".pdf"

        # ── Merge ─────────────────────────────────
        try:
            result = merge_pdfs(
                sources=files,
                add_bookmarks=add_bookmarks,
                add_page_numbers=add_page_numbers,
                password=password,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        # ── Output: JSON (metadata only) ───────────
        if output == "json":
            return Response(
                {
                    "total_pages": result["total_pages"],
                    "files_merged": result["files_merged"],
                    "size_kb": result["size_kb"],
                    "size_mb": result["size_mb"],
                    "bookmarks": result["bookmarks"],
                    "files": result["files"],
                }
            )

        # ── Output: file download (default) ────────
        response = HttpResponse(
            result["bytes"],
            content_type="application/pdf",
        )
        response["Content-Disposition"] = f'attachment; filename="{output_filename}"'
        response["Content-Length"] = len(result["bytes"])
        response["X-Total-Pages"] = result["total_pages"]
        response["X-Files-Merged"] = result["files_merged"]
        response["X-Size-KB"] = result["size_kb"]
        return response
    


class SplitPDFView(APIView):
    """
    POST /api/tools/pdf-split/
    Upload a PDF → returns split PDF files as .zip.

    Form fields:
        file        : PDF file                          (required)
        split_by    : page | chunk | range | pages      (default: page)
        chunk_size  : pages per chunk                   (default: 1)
        pages       : JSON array for pages mode         (optional)
                      e.g. [1,3,5]
        ranges      : JSON array for range mode         (optional)
                      e.g. [[1,3],[4,6]]
        password    : PDF password if encrypted         (optional)
        output      : zip | json                        (default: zip)
    """

    parser_classes = [MultiPartParser]
    permission_classes = [AllowAny]

    def post(self, request):
        file = request.FILES.get("file")
        split_by = request.data.get("split_by", "page")
        chunk_size = request.data.get("chunk_size", "1")
        pages_raw = request.data.get("pages", None)
        ranges_raw = request.data.get("ranges", None)
        password = request.data.get("password", None)
        output = request.data.get("output", "zip")

        # ── Validate ──────────────────────────────
        if not file:
            return Response(
                {"error": "No file provided."},
                status=400,
            )

        if not file.name.lower().endswith(".pdf"):
            return Response(
                {"error": "Only .pdf files are accepted."},
                status=400,
            )

        if split_by not in ("page", "chunk", "range", "pages"):
            return Response(
                {"error": "split_by must be: page, chunk, range, or pages."},
                status=400,
            )

        if output not in ("zip", "json"):
            return Response(
                {"error": "output must be: zip or json."},
                status=400,
            )

        # ── Parse chunk_size ──────────────────────
        try:
            chunk_size = int(chunk_size)
            if chunk_size < 1:
                raise ValueError
        except (ValueError, TypeError):
            return Response(
                {"error": "chunk_size must be a positive integer."},
                status=400,
            )

        # ── Parse pages ───────────────────────────
        import json as _json

        parsed_pages = None
        if pages_raw:
            try:
                if isinstance(pages_raw, str):
                    parsed_pages = _json.loads(pages_raw)
                elif isinstance(pages_raw, list):
                    parsed_pages = pages_raw
                parsed_pages = [int(p) for p in parsed_pages]
            except Exception:
                return Response(
                    {"error": "pages must be valid JSON array: [1,3,5]"},
                    status=400,
                )

        # ── Parse ranges ──────────────────────────
        parsed_ranges = None
        if ranges_raw:
            try:
                if isinstance(ranges_raw, str):
                    parsed_ranges = _json.loads(ranges_raw)
                elif isinstance(ranges_raw, list):
                    parsed_ranges = ranges_raw
            except Exception:
                return Response(
                    {"error": "ranges must be valid JSON array: [[1,3],[4,6]]"},
                    status=400,
                )

        # ── Validate required params ───────────────
        if split_by == "range" and not parsed_ranges:
            return Response(
                {"error": "ranges is required for range mode. e.g. [[1,3],[4,6]]"},
                status=400,
            )

        if split_by == "pages" and not parsed_pages:
            return Response(
                {"error": "pages is required for pages mode. e.g. [1,3,5]"},
                status=400,
            )

        # ── Split ─────────────────────────────────
        try:
            parts = split_pdf(
                file,
                split_by=split_by,
                pages=parsed_pages,
                chunk_size=chunk_size,
                ranges=parsed_ranges,
                password=password,
            )
        except ValueError as e:
            return Response({"error": str(e)}, status=400)
        except Exception as e:
            return Response({"error": str(e)}, status=500)

        base_name = file.name.replace(".pdf", "").replace(".PDF", "")

        # ── Output: JSON (metadata only) ──────────
        if output == "json":
            return Response(
                {
                    "total_parts": len(parts),
                    "split_by": split_by,
                    "parts": [
                        {
                            "index": p["index"],
                            "filename": p["filename"],
                            "pages": p["pages"],
                            "start_page": p["start_page"],
                            "end_page": p["end_page"],
                            "size_kb": p["size_kb"],
                        }
                        for p in parts
                    ],
                }
            )

        # ── Output: ZIP of all parts ───────────────
        import zipfile

        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for part in parts:
                zf.writestr(part["filename"], part["bytes"])

        zip_buffer.seek(0)

        response = HttpResponse(
            zip_buffer.read(),
            content_type="application/zip",
        )
        response["Content-Disposition"] = (
            f'attachment; filename="{base_name}_split.zip"'
        )
        response["X-Total-Parts"] = len(parts)
        response["X-Split-By"] = split_by
        return response

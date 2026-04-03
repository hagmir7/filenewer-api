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

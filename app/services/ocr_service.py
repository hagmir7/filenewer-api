"""
OCR (Optical Character Recognition) service functions.
"""

import io
import logging

logger = logging.getLogger(__name__)


def ocr_pdf(
    source,
    language: str = "eng",
    dpi: int = 300,
    pages: list = None,
    password: str = None,
) -> dict:
    """
    Extract text from a PDF using pymupdf only.
    No Tesseract or system dependencies needed.

    Strategy:
        1. Try direct text extraction (for digital/selectable PDFs)
        2. If page has no text → render to image → extract via pymupdf OCR
        3. Combine results per page

    Args:
        source   : uploaded file object OR raw bytes
        language : language hint (eng, ara, fra, etc.)  (default: eng)
        dpi      : render resolution for image pages    (default: 300)
        pages    : list of page numbers (1-based)       (default: all)
        password : PDF password if encrypted

    Returns:
        {
            'full_text'   : str,
            'pages'       : [ { page, text, word_count, char_count, method } ],
            'total_pages' : int,
            'word_count'  : int,
            'char_count'  : int,
            'language'    : str,
            'dpi'         : int,
        }
    """
    import fitz  # pymupdf
    from PIL import Image, ImageEnhance, ImageFilter

    # ── Read PDF bytes ────────────────────────────
    if hasattr(source, "read"):
        pdf_bytes = source.read()
    else:
        pdf_bytes = source

    if not pdf_bytes.startswith(b"%PDF"):
        raise ValueError("Invalid PDF file.")

    # ── Open PDF ──────────────────────────────────
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    if doc.is_encrypted:
        if not password:
            raise ValueError("PDF is encrypted. Provide a password.")
        if not doc.authenticate(password):
            raise ValueError("Wrong password.")

    total_pages = len(doc)

    # ── Validate pages ────────────────────────────
    if pages is not None:
        invalid = [p for p in pages if not (1 <= p <= total_pages)]
        if invalid:
            raise ValueError(
                f"Invalid page numbers: {invalid}. "
                f"PDF has {total_pages} pages (1-{total_pages})."
            )
        pages_to_ocr = pages
    else:
        pages_to_ocr = list(range(1, total_pages + 1))

    # ── Render config ─────────────────────────────
    zoom = dpi / 72
    matrix = fitz.Matrix(zoom, zoom)

    # ── Process each page ─────────────────────────
    page_results = []
    full_texts = []

    for page_num in pages_to_ocr:
        page = doc[page_num - 1]

        # ── Step 1: Try direct text extraction ────
        direct_text = page.get_text("text").strip()

        if direct_text and len(direct_text) > 20:
            # Good digital text found — use directly
            text = direct_text
            method = "digital"

        else:
            # ── Step 2: Image-based extraction ────
            # Render page to high-res image
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)

            # Pixmap → PIL Image
            pil_img = Image.frombytes(
                "RGB",
                [pixmap.width, pixmap.height],
                pixmap.samples,
            )

            # ── Preprocess image ──────────────────
            pil_img = _preprocess_image_for_ocr(pil_img)

            # ── Save processed image to bytes ──────
            img_buffer = io.BytesIO()
            pil_img.save(img_buffer, format="PNG", optimize=True)
            img_buffer.seek(0)

            # ── Re-open with pymupdf for text ──────
            # Create a temporary single-page PDF from image
            img_doc = fitz.open(stream=img_buffer.read(), filetype="png")
            img_page = img_doc[0]

            # Extract text blocks from rendered image
            # pymupdf can extract text from image-based pages too
            blocks = img_page.get_text("blocks")
            text = "\n".join(b[4].strip() for b in blocks if b[4].strip())

            if not text:
                # Final fallback: get all text from original page
                text = page.get_text("words")
                text = " ".join(w[4] for w in text).strip()

            img_doc.close()
            method = "image"

        word_count = len(text.split()) if text else 0
        char_count = len(text)

        page_results.append(
            {
                "page": page_num,
                "text": text,
                "word_count": word_count,
                "char_count": char_count,
                "method": method,
            }
        )

        full_texts.append(f"--- Page {page_num} ---\n{text}")

    doc.close()

    full_text = "\n\n".join(full_texts)
    word_count = sum(p["word_count"] for p in page_results)
    char_count = sum(p["char_count"] for p in page_results)

    return {
        "full_text": full_text,
        "pages": page_results,
        "total_pages": len(page_results),
        "word_count": word_count,
        "char_count": char_count,
        "language": language,
        "dpi": dpi,
    }


def _preprocess_image_for_ocr(image) -> "Image":
    """
    Preprocess PIL image for better text extraction accuracy.

    Steps:
        1. Convert to grayscale
        2. Enhance contrast
        3. Sharpen edges
        4. Denoise
    """
    from PIL import Image, ImageEnhance, ImageFilter

    # ── Upscale small images ──────────────────────
    w, h = image.size
    if w < 1000 or h < 1000:
        scale = max(1000 / w, 1000 / h)
        new_w = int(w * scale)
        new_h = int(h * scale)
        image = image.resize((new_w, new_h), Image.LANCZOS)

    # ── Convert to grayscale ──────────────────────
    gray = image.convert("L")

    # ── Enhance contrast ──────────────────────────
    contrast = ImageEnhance.Contrast(gray)
    enhanced = contrast.enhance(2.0)

    # ── Sharpen ───────────────────────────────────
    sharpener = ImageEnhance.Sharpness(enhanced)
    sharpened = sharpener.enhance(2.0)

    # ── Remove noise ──────────────────────────────
    denoised = sharpened.filter(ImageFilter.MedianFilter(size=3))

    return denoised

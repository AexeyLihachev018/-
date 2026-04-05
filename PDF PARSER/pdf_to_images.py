import base64
import io
import fitz  # PyMuPDF


def pdf_to_base64_images(pdf_path: str, dpi: int = 150) -> list[str]:
    """Конвертирует страницы PDF в список base64-encoded PNG (через PyMuPDF, без poppler)."""
    doc = fitz.open(pdf_path)
    zoom = dpi / 72  # PDF базовый DPI = 72
    matrix = fitz.Matrix(zoom, zoom)

    result = []
    for page in doc:
        pix = page.get_pixmap(matrix=matrix)
        png_bytes = pix.tobytes("png")
        b64 = base64.b64encode(png_bytes).decode("utf-8")
        result.append(b64)

    doc.close()
    return result

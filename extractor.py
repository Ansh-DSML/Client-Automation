import io
import pdfplumber


def extract_text(pdf_bytes: bytes) -> str:
    """
    Extracts all text from a PDF given its raw bytes.
    Returns a single string with all pages joined by newline.
    pdfplumber preserves spatial layout better than pypdf for
    structured documents like BEL bid invitations.
    """
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        pages = [page.extract_text() or "" for page in pdf.pages]
    return "\n".join(pages)

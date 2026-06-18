import io
from supabase import create_client, Client
from config import SUPABASE_URL, SUPABASE_SERVICE_KEY, SUPABASE_BUCKET

_client: Client = None

def get_client() -> Client:
    global _client
    if _client is None:
        _client = create_client(SUPABASE_URL, SUPABASE_SERVICE_KEY)
    return _client


def upload_pdf(pdf_bytes: bytes, filename: str) -> str:
    """
    Uploads PDF bytes to Supabase storage.
    Returns the storage path (used to download later).
    Raises an exception if upload fails.
    """
    client = get_client()
    # Storage path inside the bucket: rfq-pdfs/BID7000630957.PDF
    storage_path = filename

    client.storage.from_(SUPABASE_BUCKET).upload(
        path=storage_path,
        file=pdf_bytes,
        file_options={"content-type": "application/pdf", "upsert": "true"}
    )
    return storage_path


def log_upload_to_db(filename: str, storage_path: str) -> str:
    """
    Inserts a record into rfq_uploads table.
    Returns the UUID of the inserted row.
    """
    client = get_client()
    result = client.table("rfq_uploads").insert({
        "filename": filename,
        "storage_path": storage_path,
    }).execute()
    return result.data[0]["id"]


def download_pdf(storage_path: str) -> bytes:
    """
    Downloads a PDF from Supabase storage by its storage path.
    Returns raw bytes.
    """
    client = get_client()
    response = client.storage.from_(SUPABASE_BUCKET).download(storage_path)
    return response


def mark_processed(record_id: str, company: str, row_count: int):
    """
    Updates the rfq_uploads row after successful processing.
    """
    client = get_client()
    client.table("rfq_uploads").update({
        "processed": True,
        "company_detected": company,
        "row_count": row_count,
    }).eq("id", record_id).execute()


def get_unprocessed() -> list[dict]:
    """
    Returns all rows in rfq_uploads where processed = False.
    Used if you want to reprocess failed PDFs later.
    """
    client = get_client()
    result = client.table("rfq_uploads").select("*").eq("processed", False).execute()
    return result.data

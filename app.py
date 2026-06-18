import streamlit as st
from supabase_client import upload_pdf, log_upload_to_db, download_pdf, mark_processed
from extractor import extract_text
from groq_detector import detect_company
from router import route
from excel_writer import append_rows
import gem_excel_writer
from config import OUTPUT_EXCEL_PATH, GEM_OUTPUT_EXCEL_PATH

st.set_page_config(page_title="RFQ Processor", page_icon="📋", layout="centered")

st.markdown("""
<style>
div.stButton > button[kind="primary"],
div.stDownloadButton > button[kind="primary"] {
    background-color: #0E577C !important;
    border-color: #0E577C !important;
    color: white !important;
}
div.stButton > button[kind="primary"]:hover,
div.stDownloadButton > button[kind="primary"]:hover {
    background-color: #0a3f5a !important;
    border-color: #0a3f5a !important;
}
</style>
""", unsafe_allow_html=True)

st.title("RFQ PDF Processor")
st.caption("Upload one or more RFQ PDFs. They will be stored in Supabase and processed into Excel.")

uploaded_files = st.file_uploader(
    "Upload RFQ PDFs",
    type="pdf",
    accept_multiple_files=True,
)

if uploaded_files:
    st.write(f"**{len(uploaded_files)} file(s) selected**")

    if st.button("Process PDFs", type="primary", use_container_width=True):

        progress = st.progress(0)
        total    = len(uploaded_files)
        results  = []

        for idx, uf in enumerate(uploaded_files):
            st.markdown(f"---\n**Processing: {uf.name}**")

            try:
                pdf_bytes = uf.read()

                # ── Step 1: Upload to Supabase storage ──────────────────────
                with st.spinner("Uploading to Supabase..."):
                    storage_path = upload_pdf(pdf_bytes, uf.name)
                    record_id    = log_upload_to_db(uf.name, storage_path)
                st.success(f"Stored in Supabase → `{storage_path}`")

                # ── Step 2: Extract full text ────────────────────────────────
                with st.spinner("Extracting text from PDF..."):
                    full_text = extract_text(pdf_bytes)
                st.success(f"Text extracted — {len(full_text):,} characters")

                # ── Step 3: Detect company via Groq ─────────────────────────
                with st.spinner("Detecting company via Groq..."):
                    detection = detect_company(full_text)

                company_code = detection["company_code"]
                company_name = detection["company_full_name"]
                confidence   = detection["confidence"]

                if confidence == "low":
                    st.warning(
                        f"Company detected as **{company_name}** ({company_code}) "
                        f"but confidence is LOW. Verify manually."
                    )
                else:
                    st.success(f"Company: **{company_name}** ({company_code}) — {confidence} confidence")

                # ── Step 4: Route to parser ──────────────────────────────────
                with st.spinner(f"Running {company_code} parser..."):
                    rows = route(full_text, company_code, pdf_bytes=pdf_bytes)
                st.success(f"Parsed → {len(rows)} row(s)")

                # ── Step 5: Write to Excel ───────────────────────────────────
                with st.spinner("Writing to Excel..."):
                    if company_code == "GEM":
                        saved_path = gem_excel_writer.append_rows(rows)
                    else:
                        saved_path = append_rows(rows)
                st.success(f"Saved to `{saved_path}`")

                # ── Step 6: Update Supabase record ───────────────────────────
                mark_processed(record_id, company_code, len(rows))

                results.append({
                    "file": uf.name,
                    "company": f"{company_name} ({company_code})",
                    "rows": len(rows),
                    "status": "Done",
                })

            except Exception as e:
                st.error(f"Failed on {uf.name}: {e}")
                results.append({
                    "file": uf.name,
                    "company": "Error",
                    "rows": 0,
                    "status": str(e),
                })

            progress.progress((idx + 1) / total)

        # ── Summary table ────────────────────────────────────────────────────
        if results:
            st.markdown("---")
            st.subheader("Summary")
            import pandas as pd
            st.dataframe(pd.DataFrame(results), use_container_width=True)
            st.info(f"Excel file updated at: `{saved_path}`")

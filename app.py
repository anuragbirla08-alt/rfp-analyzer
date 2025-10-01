"""Streamlit application for RFP and SOW analysis using Google Gemini models."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from io import BytesIO
from typing import Dict, Iterable, List, Set, Tuple
from urllib.parse import unquote, urlparse

import fitz
import pandas as pd
import requests
import streamlit as st
from docx import Document
from google import genai


def _initial_api_key() -> str:
    """Fetch a default Gemini API key from Streamlit secrets or environment."""

    try:
        candidate = st.secrets.get("GEMINI_API_KEY", "")  # type: ignore[attr-defined]
    except Exception:
        candidate = ""
    return candidate or os.getenv("GEMINI_API_KEY", "")


DEFAULT_API_KEY = _initial_api_key()
DOWNLOAD_TIMEOUT = 20
MAX_LINKED_DOCUMENTS = 8
MAX_CHARS = 60000


@dataclass
class LinkedDocument:
    """Container for downloaded linked documents."""

    url: str
    filename: str
    media_type: str
    text: str
    data: bytes

# ---------------------------------------------------------------------------
# Page configuration and sidebar controls
# ---------------------------------------------------------------------------

st.set_page_config(page_title="RFP Analysis Tool", page_icon="📄", layout="wide")

st.title("📄 RFP & SOW Analysis Automation Tool")
st.write(
    "Upload government tender PDFs (RFP or SOW documents) to automatically "
    "extract key information."
)

if "api_key_input" not in st.session_state:
    st.session_state["api_key_input"] = DEFAULT_API_KEY

with st.sidebar:
    st.header("⚙️ Configuration")
    if DEFAULT_API_KEY:
        st.success("Using API key from Streamlit secrets or environment by default.")
    st.text_input(
        "Gemini API Key",
        type="password",
        key="api_key_input",
        help="Persist a key in Streamlit secrets or the GEMINI_API_KEY environment variable to avoid re-entering it.",
    )
    api_key = st.session_state.get("api_key_input", "").strip()
    st.caption(
        "Don't have a key? [Get one from Google AI Studio](https://aistudio.google.com/)"
    )

    st.divider()

    st.subheader("Model Selection")
    model_choice = st.radio(
        "Choose AI Model:",
        ("gemini-2.5-flash (Faster)", "gemini-2.5-pro (More Accurate)"),
        help="Flash is faster but Pro gives better extraction quality",
    )
    SELECTED_MODEL = "gemini-2.5-flash" if "flash" in model_choice else "gemini-2.5-pro"

    st.divider()

    st.subheader("📋 Extracted Fields")
    with st.expander("RFP Fields"):
        st.markdown(
            """
            - Procurement basics (Bid ID, item category, entity, key dates)
            - Financials (estimated value, turnover, solvency, EMD, BG)
            - Experience & compliance (years, certifications, labour laws, OEM clauses)
            - Delivery (contract period, locations, SLA, penalties, payment terms)
            - Teaming (consortium/MSE/startup provisions, subcontracting limits)
            - Evaluation (technical/financial weighting, past performance scoring)
            - Risk (liquidated damages, termination, arbitration, statutory liabilities)
            - Linked resources (corrigenda, SOW, ATC, drawings, annexures)
            """
        )

    with st.expander("SOW Fields"):
        st.markdown(
            """
            - Technical architecture, integrations, and compliance norms
            - Deliverables, milestones, acceptance gates, and documentation
            - Service levels, KPIs, governance cadence, escalation paths
            - Resource mix, staffing qualifications, onboarding requirements
            - Implementation & transition methodology, change management
            - Warranty, support, training, and knowledge transfer expectations
            """
        )

# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------


def _safe_filename(url: str) -> str:
    parsed = urlparse(url)
    if parsed.path:
        candidate = os.path.basename(parsed.path)
        if candidate:
            return unquote(candidate)
    return unquote(parsed.netloc or "linked_document")


def _extract_text_from_pdf_bytes(data: bytes) -> str:
    document = fitz.open(stream=data, filetype="pdf")
    try:
        return "".join(page.get_text() for page in document)
    finally:
        document.close()


def _extract_text_from_docx_bytes(data: bytes) -> str:
    doc = Document(BytesIO(data))
    paragraphs = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    return "\n".join(paragraphs)


def download_linked_documents(urls: Iterable[str]) -> List[LinkedDocument]:
    """Fetch linked documents (PDF/DOCX) referenced inside the tender."""

    downloaded: List[LinkedDocument] = []
    for url in urls:
        if len(downloaded) >= MAX_LINKED_DOCUMENTS:
            st.info(
                f"Reached the download limit of {MAX_LINKED_DOCUMENTS} linked documents; skip remaining URLs."
            )
            break
        try:
            response = requests.get(
                url,
                timeout=DOWNLOAD_TIMEOUT,
                headers={
                    "User-Agent": "Mozilla/5.0 (Gemini-RFP-Analyzer)",
                },
            )
            response.raise_for_status()
        except Exception as exc:
            st.warning(f"Unable to download linked document {url}: {exc}")
            continue

        data = response.content
        content_type = response.headers.get("Content-Type", "").lower()
        filename = _safe_filename(url)

        text = ""
        media_type = "unknown"
        try:
            if "pdf" in content_type or filename.lower().endswith(".pdf"):
                text = _extract_text_from_pdf_bytes(data)
                media_type = "pdf"
            elif any(ext in content_type for ext in ("word", "docx")) or filename.lower().endswith(
                (".docx",)
            ):
                text = _extract_text_from_docx_bytes(data)
                media_type = "docx"
            else:
                media_type = content_type or "unknown"
        except Exception as exc:  # pragma: no cover - dependent on remote docs
            st.warning(f"Unable to parse linked document {filename}: {exc}")
            continue

        if not text.strip():
            st.info(
                f"Downloaded linked document {filename} but could not extract text; skipping it in analysis."
            )
            continue

        downloaded.append(
            LinkedDocument(url=url, filename=filename, media_type=media_type, text=text, data=data)
        )

    return downloaded


@st.cache_resource(show_spinner=False)
def get_client(api_key_value: str) -> genai.Client | None:
    """Return a cached Gemini client instance for the provided API key."""
    if not api_key_value:
        return None
    try:
        return genai.Client(api_key=api_key_value)
    except Exception as exc:  # pragma: no cover - protective UI feedback
        st.error(f"Failed to initialise Gemini client: {exc}")
        return None


def chunk_text(text: str, *, chunk_size: int = MAX_CHARS, overlap: int = 500) -> List[str]:
    """Split long text into overlapping chunks suitable for model ingestion."""
    if len(text) <= chunk_size:
        return [text]

    chunks: List[str] = []
    start = 0
    while start < len(text):
        end = min(start + chunk_size, len(text))
        chunks.append(text[start:end])
        if end == len(text):
            break
        start = max(0, end - overlap)
    return chunks


def extract_text_and_urls_from_pdf(pdf_file) -> Tuple[str | None, List[str]]:
    """Extract text content and hyperlinks from a PDF file."""
    try:
        pdf_bytes = pdf_file.read()
        document = fitz.open(stream=pdf_bytes, filetype="pdf")
        text = _extract_text_from_pdf_bytes(pdf_bytes)

        urls: List[str] = []
        for page in document:
            for link in page.get_links():
                uri = link.get("uri")
                if uri and uri.startswith(("http://", "https://")) and uri not in urls:
                    urls.append(uri)

        document.close()
        return text, urls
    except Exception as exc:
        st.error(f"Error reading PDF {pdf_file.name}: {exc}")
        return None, []


def _build_rfp_prompt(
    text: str,
    filename: str,
    urls: Iterable[str],
    linked_docs: Iterable[LinkedDocument],
) -> str:
    urls_string = "; ".join(urls) if urls else "No linked documents found"
    linked_doc_overview = "\n".join(
        f"- {doc.filename} ({doc.media_type}) from {doc.url}" for doc in linked_docs
    )
    linked_doc_block = (
        "Linked artefacts referenced in the tender:\n" + linked_doc_overview
        if linked_doc_overview
        else "No additional linked artefacts were available."
    )
    return f"""
You are an expert Government Tender Analyst. Review the tender corpus below (including annexures downloaded from links) and convert every relevant insight into a JSON object for bid/no-bid decision making.

Tender Document: {filename}
{linked_doc_block}
---
{text}
---

Instructions:
1. Populate every field even if the detail appears only in annexures or ATC documents. Use "Not Found" if the tender is silent.
2. Keep values concise but information-rich. Normalise currencies to INR Lakhs/Crores and durations to months where possible.
3. Summarise eligibility criteria in actionable language referencing turnover, experience, certifications, OEM mandates, manpower, etc.
4. Extract risk, payment, and evaluation nuances that impact delivery strategy or commercial positioning.
5. Retain URLs exactly as provided for traceability.

JSON Schema (use these exact keys):
{{
  "File_Name": "{filename}",
  "Tender_Title": "Official tender title or short name",
  "Bid_Number": "e.g., GEM/2025/B/6530170",
  "Procurement_Category": "Goods/Services/Works",
  "Item_Category": "Core service/item being procured",
  "Ministry_Department_Office": "Full name of procuring entity",
  "Buyer_Contact_Details": "Name, email, phone if provided",
  "Bid_End_Date": "Date/time bid closes",
  "Bid_Opening_Date": "Date/time technical bid opens",
  "Pre_Bid_Meeting_Details": "Date/time/location or virtual link",
  "Contract_Period_Months": "Duration in months",
  "Project_Location": "Delivery locations / site spread",
  "Estimated_Value_INR": "Estimated value in Lakhs/Crores",
  "Budgetary_Source": "Budget head / funding info",
  "Min_Annual_Turnover_Req_INR": "Average annual turnover requirement",
  "Net_Worth_or_Solvency_Req": "Net worth / solvency / credit lines",
  "Years_Past_Experience_Req": "Years of similar experience required",
  "Similar_Work_Definition": "How 'similar work' is defined",
  "Required_Certifications": "Mandatory registrations/certifications",
  "Key_Personnel_Qualifications": "Minimum qualifications / certifications for staff",
  "Consortium_or_JV_Allowed": "Yes/No with conditions",
  "Subcontracting_Limits": "Permissible subcontracting percentage",
  "MSE_Exemption_Available": "Yes/No plus scope",
  "Startup_Exemption_Available": "Yes/No plus scope",
  "EMD_Required_Amount": "EMD amount or exemption",
  "Bid_Security_Options": "BG/online payment/waivers",
  "Performance_Bank_Guarantee": "PBG % and validity",
  "Advance_Payment_Terms": "Any mobilisation advance conditions",
  "Payment_Terms": "Milestone / monthly / completion payment pattern",
  "Penalties_LD_Clauses": "Liquidated damages / penalty triggers",
  "SLA_and_Service_Credits": "Key SLA metrics and service credits",
  "Scope_of_Work_Summary": "1-2 paragraph scope summary",
  "Deliverables_and_Milestones": "List of key deliverables/milestones",
  "Evaluation_Methodology": "QCBS/L1/combined scoring etc.",
  "Technical_to_Financial_Weightage": "Weight split if QCBS",
  "Go_NoGo_Eligibility_Summary": "Action-oriented summary of minimum qualifications",
  "Critical_Risks_Notes": "Single paragraph on major risks/contractual burdens",
  "Additional_Observations": "Other insights influencing bid decision",
  "Key_Linked_Documents_URLs": "{urls_string}"
}}

Return ONLY valid JSON without markdown formatting.
"""


def _build_sow_prompt(
    text: str,
    filename: str,
    linked_docs: Iterable[LinkedDocument],
) -> str:
    linked_doc_overview = "\n".join(
        f"- {doc.filename} ({doc.media_type}) from {doc.url}" for doc in linked_docs
    )
    linked_doc_block = (
        "Linked artefacts referenced in the tender:\n" + linked_doc_overview
        if linked_doc_overview
        else "No additional linked artefacts were available."
    )
    return f"""
You are an expert Government Tender Analyst. Review the SOW corpus below (including annexures downloaded from links) and extract exhaustive delivery expectations into JSON.

Document: {filename}
{linked_doc_block}
---
{text}
---

Instructions:
1. Capture technical, operational, governance, and transition requirements in depth.
2. Use "Not Found" when the requirement is absent.
3. Highlight resourcing, tooling, and acceptance controls that influence execution.

JSON Schema (use these exact keys):
{{
  "File_Name": "{filename}",
  "Bid_Number": "Reference/Bid number if mentioned",
  "Document_Type": "SOW",
  "Workstreams": "Major workstreams/modules",
  "Technical_Requirements": "Detailed technical specifications",
  "Integrations_and_Dependencies": "Systems to integrate / prerequisites",
  "Deliverables": "Expected deliverables list",
  "Timeline_Schedule": "Project timeline and milestones",
  "Resource_Profile": "Required roles/skillsets/onsite-offsite split",
  "Tools_and_Technologies": "Preferred/mandated tools and tech stack",
  "Performance_Metrics": "KPIs and success criteria",
  "Quality_Standards": "Quality requirements and certifications",
  "Compliance_Requirements": "Regulatory/legal/security compliance",
  "Implementation_Methodology": "Required approach/methodology",
  "Change_Management_and_Governance": "Meetings, reporting, approvals",
  "Support_and_Maintenance": "Post-implementation support expectations",
  "Training_Requirements": "Training / KT obligations",
  "Warranty_and_Penalties": "Warranty/support commitments and penalties",
  "Acceptance_Criteria": "Project acceptance criteria",
  "Assumptions_and_Exclusions": "Any stated assumptions or exclusions"
}}

Return ONLY valid JSON without markdown formatting.
"""


def _response_to_text(response) -> str:
    """Normalise Gemini responses to raw text."""

    if hasattr(response, "text") and response.text:
        return response.text

    candidates = getattr(response, "candidates", None)
    if candidates:
        for candidate in candidates:
            content = getattr(candidate, "content", None)
            if not content:
                continue
            parts = getattr(content, "parts", None)
            if not parts:
                continue
            for part in parts:
                text = getattr(part, "text", None)
                if text:
                    return text
    raise ValueError("No text content in Gemini response")


def call_gemini_text(prompt: str, *, client: genai.Client, **generation_kwargs) -> str:
    """Invoke Gemini and return plain text output."""

    try:
        response = client.models.generate_content(
            model=SELECTED_MODEL,
            contents=prompt,
            generation_config=generation_kwargs or None,
        )
        return _response_to_text(response).strip()
    except Exception as exc:  # pragma: no cover - network/external
        st.warning(f"Gemini text generation failed: {exc}")
        return ""


def _clean_json_text(raw_text: str) -> str:
    """Remove markdown fences and whitespace around a JSON payload."""

    text = raw_text.strip()
    if text.startswith("```"):
        # Split once to drop the opening fence and keep the remainder.
        text = text.split("\n", 1)[-1]
    if text.endswith("```"):
        text = text.rsplit("\n", 1)[0]
    return text.strip()


def _coerce_json_from_text(raw_text: str) -> Dict[str, str] | None:
    """Best-effort conversion of an LLM response into JSON."""

    text = _clean_json_text(raw_text)
    if not text:
        return None

    try:
        return json.loads(text)
    except json.JSONDecodeError:
        pass

    # Attempt to locate the first JSON object within the response body.
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        candidate = text[start : end + 1]
        try:
            return json.loads(candidate)
        except json.JSONDecodeError:
            return None
    return None


def call_gemini_json(prompt: str, *, client: genai.Client) -> Dict[str, str] | None:
    """Invoke the Gemini model and coerce the response into JSON."""

    def _run_request(**generation_kwargs) -> Dict[str, str] | None:
        response = client.models.generate_content(
            model=SELECTED_MODEL,
            contents=prompt,
            generation_config=generation_kwargs or None,
        )
        text = _response_to_text(response)
        return _coerce_json_from_text(text)

    try:
        # First, request native JSON. Fall back to free-form text if the model
        # emits validation errors or extra commentary.
        data = _run_request(response_mime_type="application/json")
        if data is not None:
            return data

        data = _run_request()
        if data is not None:
            return data

        st.error("Failed to parse AI response as JSON; the model did not return structured data.")
        return None
    except Exception as exc:  # pragma: no cover - defensive UI feedback
        st.error(f"AI Error: {exc}")
        return None


def reduce_text_for_prompt(text: str, *, client: genai.Client) -> str:
    """Compress lengthy tender text into a model-friendly summary."""

    chunks = chunk_text(text)
    if len(chunks) == 1 and len(chunks[0]) <= MAX_CHARS:
        return chunks[0]

    summaries: List[str] = []
    for idx, chunk in enumerate(chunks, start=1):
        summary_prompt = (
            "Summarise the following tender excerpt into crisp bullet points focusing on "
            "eligibility criteria, deliverables, risks, financials, and any referenced annexures.\n\n"
            f"Excerpt {idx}:\n{chunk}\n\nReturn 5-8 bullets capturing all unique obligations."
        )
        summary = call_gemini_text(summary_prompt, client=client, max_output_tokens=512)
        if summary:
            summaries.append(f"Excerpt {idx} summary:\n{summary}")

    if summaries:
        condensed = "\n\n".join(summaries)
        return condensed[:MAX_CHARS]

    return text[:MAX_CHARS]


def prepare_analysis_text(
    primary_text: str, linked_docs: Iterable[LinkedDocument], *, client: genai.Client
) -> str:
    """Combine primary PDF content with linked annexures and compress if required."""

    sections = [primary_text.strip()]
    for doc in linked_docs:
        if doc.text.strip():
            sections.append(
                f"Linked Document: {doc.filename} ({doc.media_type}) from {doc.url}\n{doc.text.strip()}"
            )

    combined = "\n\n".join(section for section in sections if section)
    if not combined:
        return ""

    if len(combined) <= MAX_CHARS:
        return combined

    return reduce_text_for_prompt(combined, client=client)


MISSING_STRINGS: Set[str] = {"", "na", "n/a", "not found", "nil", "none", "not applicable"}


def value_present(value: str | None) -> bool:
    if value is None:
        return False
    normalised = value.strip().lower()
    return normalised not in MISSING_STRINGS and normalised != "-"


def build_field_coverage(records: List[Dict[str, str]]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame()

    fields: List[str] = []
    for record in records:
        for key in record.keys():
            if key not in ("File_Name", "Document_Type") and key not in fields:
                fields.append(key)

    coverage_rows: List[Dict[str, object]] = []
    for field in fields:
        present_count = 0
        example = ""
        for record in records:
            value = record.get(field)
            if value_present(value):
                present_count += 1
                if not example:
                    example = f"{record.get('File_Name', 'Unknown')}: {value}"
        coverage_rows.append(
            {
                "Field": field,
                "Documents_with_data": present_count,
                "Documents_missing_data": len(records) - present_count,
                "Example_Value": example,
            }
        )

    return pd.DataFrame(coverage_rows)


def build_requirement_matrix(records: List[Dict[str, str]]) -> pd.DataFrame:
    if not records:
        return pd.DataFrame()

    matrix = pd.DataFrame(records)
    if "File_Name" in matrix.columns:
        matrix = matrix.set_index("File_Name")

    return matrix.transpose()


def extract_rfp_data(
    text: str,
    filename: str,
    urls: Iterable[str],
    linked_docs: Iterable[LinkedDocument],
    *,
    client: genai.Client,
):
    context = prepare_analysis_text(text, linked_docs, client=client)
    prompt = _build_rfp_prompt(context, filename, urls, linked_docs)
    return call_gemini_json(prompt, client=client)


def extract_sow_data(
    text: str,
    filename: str,
    linked_docs: Iterable[LinkedDocument],
    *,
    client: genai.Client,
):
    context = prepare_analysis_text(text, linked_docs, client=client)
    prompt = _build_sow_prompt(context, filename, linked_docs)
    data = call_gemini_json(prompt, client=client)
    if data is not None:
        data.setdefault("Document_Type", "SOW")
    return data


# ---------------------------------------------------------------------------
# Main UI logic
# ---------------------------------------------------------------------------

uploaded_files = st.file_uploader(
    "📁 Choose PDF files",
    type="pdf",
    accept_multiple_files=True,
    help="Upload one or more RFP or SOW documents",
)

if uploaded_files:
    st.subheader("Document Type Selection")
    st.caption("Select the document type for each uploaded file")

    document_types: Dict[str, str] = {}
    columns = st.columns(min(len(uploaded_files), 3))
    for index, file in enumerate(uploaded_files):
        with columns[index % len(columns)]:
            document_types[file.name] = st.radio(
                label=file.name,
                options=("RFP", "SOW"),
                key=f"doc_type_{index}",
                horizontal=True,
            )
else:
    document_types = {}

analyze_disabled = not uploaded_files

if st.button("🚀 Analyze Documents", type="primary", disabled=analyze_disabled):
    if not api_key:
        st.warning("⚠️ Please enter your Gemini API Key in the sidebar to proceed.")
    elif not uploaded_files:
        st.warning("⚠️ Please upload at least one PDF file.")
    else:
        client = get_client(api_key)
        if client is None:
            st.stop()

        rfp_results: List[Dict[str, str]] = []
        sow_results: List[Dict[str, str]] = []
        linked_documents: Dict[str, List[LinkedDocument]] = {}

        progress_bar = st.progress(0)
        status_placeholder = st.empty()

        for index, pdf in enumerate(uploaded_files, start=1):
            status_placeholder.info(
                f"📄 Processing: **{pdf.name}** ({index}/{len(uploaded_files)})"
            )

            text, urls = extract_text_and_urls_from_pdf(pdf)
            if not text:
                progress_bar.progress(index / len(uploaded_files))
                continue

            linked_docs = download_linked_documents(urls)
            linked_documents[pdf.name] = linked_docs

            if linked_docs:
                with st.expander(
                    f"Linked artefacts fetched for {pdf.name}", expanded=False
                ):
                    for doc_index, doc in enumerate(linked_docs, start=1):
                        st.markdown(
                            f"**{doc_index}. {doc.filename}** ({doc.media_type.upper()})  "
                            f"[Open source]({doc.url})"
                        )
                        mime = (
                            "application/pdf"
                            if doc.media_type == "pdf"
                            else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            if doc.media_type == "docx"
                            else "application/octet-stream"
                        )
                        st.download_button(
                            label=f"Download {doc.filename}",
                            data=doc.data,
                            file_name=doc.filename,
                            mime=mime,
                            key=f"download_{pdf.name}_{doc_index}",
                        )

            doc_type = document_types.get(pdf.name, "RFP")
            if doc_type == "RFP":
                result = extract_rfp_data(text, pdf.name, urls, linked_docs, client=client)
                if result:
                    rfp_results.append(result)
                    st.success(f"✅ Successfully analyzed {pdf.name} (RFP)")
            else:
                result = extract_sow_data(text, pdf.name, linked_docs, client=client)
                if result:
                    sow_results.append(result)
                    st.success(f"✅ Successfully analyzed {pdf.name} (SOW)")

            progress_bar.progress(index / len(uploaded_files))

        status_placeholder.success("✨ Analysis Complete!")

        if not (rfp_results or sow_results):
            st.error("❌ No data was successfully extracted. Please check your documents and try again.")
            st.stop()

        st.divider()
        st.header("📊 Analysis Results")

        output_buffer = BytesIO()
        json_payload: Dict[str, object] = {
            "linked_documents": {
                file_name: [
                    {
                        "filename": doc.filename,
                        "url": doc.url,
                        "media_type": doc.media_type,
                    }
                    for doc in docs
                ]
                for file_name, docs in linked_documents.items()
                if docs
            }
        }

        with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
            if rfp_results:
                df_rfp = pd.DataFrame(rfp_results)
                df_rfp.to_excel(writer, index=False, sheet_name="RFP_Data")
                json_payload["rfp"] = rfp_results

                st.subheader("RFP Documents")
                st.dataframe(df_rfp, use_container_width=True)

                if "Bid_End_Date" in df_rfp:
                    st.markdown("### 🗓️ Upcoming Bid Deadlines")
                    try:
                        deadline_series = pd.to_datetime(df_rfp["Bid_End_Date"], errors="coerce")
                        upcoming = (
                            df_rfp.assign(_deadline=deadline_series)
                            .dropna(subset=["_deadline"])
                            .sort_values("_deadline")
                            .head(5)
                        )
                        if not upcoming.empty:
                            st.dataframe(
                                upcoming[["File_Name", "Bid_Number", "Bid_End_Date", "Go_NoGo_Eligibility_Summary"]],
                                use_container_width=True,
                            )
                        else:
                            st.info("No valid bid deadlines were detected in the extracted data.")
                    except Exception as exc:  # pragma: no cover
                        st.warning(f"Unable to generate bid deadline summary: {exc}")

            if sow_results:
                df_sow = pd.DataFrame(sow_results)
                df_sow.to_excel(writer, index=False, sheet_name="SOW_Data")
                json_payload["sow"] = sow_results

                st.subheader("SOW Documents")
                st.dataframe(df_sow, use_container_width=True)

        excel_data = output_buffer.getvalue()

        st.download_button(
            label="📥 Download Complete Analysis as Excel",
            data=excel_data,
            file_name="rfp_sow_analysis_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        if rfp_results:
            coverage_df = build_field_coverage(rfp_results)
            if not coverage_df.empty:
                st.markdown("### 🔍 Cross-RFP requirement coverage")
                coverage_sorted = coverage_df.sort_values(
                    by="Documents_missing_data", ascending=False
                )
                st.dataframe(coverage_sorted, use_container_width=True)
                json_payload["rfp_field_coverage"] = coverage_sorted.to_dict(
                    orient="records"
                )

            matrix_df = build_requirement_matrix(rfp_results)
            if not matrix_df.empty:
                st.markdown("### 📑 RFP requirement matrix")
                st.dataframe(matrix_df, use_container_width=True)
                json_payload["rfp_requirement_matrix"] = matrix_df.to_dict()

        st.download_button(
            label="📤 Download Raw JSON",
            data=json.dumps(json_payload, indent=2),
            file_name="rfp_sow_analysis_results.json",
            mime="application/json",
        )

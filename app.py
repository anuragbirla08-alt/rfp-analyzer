"""Streamlit application for RFP and SOW analysis using Google Gemini models."""

from __future__ import annotations

import json
from io import BytesIO
from typing import Dict, Iterable, List, Tuple

import pandas as pd
import streamlit as st
from google import genai
import fitz

# ---------------------------------------------------------------------------
# Page configuration and sidebar controls
# ---------------------------------------------------------------------------

st.set_page_config(page_title="RFP Analysis Tool", page_icon="📄", layout="wide")

st.title("📄 RFP & SOW Analysis Automation Tool")
st.write(
    "Upload government tender PDFs (RFP or SOW documents) to automatically "
    "extract key information."
)

with st.sidebar:
    st.header("⚙️ Configuration")
    api_key = st.text_input("Enter your Google Gemini API Key", type="password")
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
            - Bid Number & Dates
            - Ministry/Department/Office
            - Estimated Value & Contract Period
            - Turnover & Experience Requirements
            - MSE/Startup Exemptions
            - Scope of Work Summary
            - Go/No-Go Eligibility Summary
            - Linked Document URLs
            """
        )

    with st.expander("SOW Fields"):
        st.markdown(
            """
            - Technical Requirements
            - Deliverables & Timeline
            - Performance Metrics
            - Quality Standards
            - Implementation Methodology
            - Support & Maintenance
            """
        )

# ---------------------------------------------------------------------------
# Helper utilities
# ---------------------------------------------------------------------------

MAX_CHARS = 60000


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
        text = "".join(page.get_text() for page in document)

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


def _build_rfp_prompt(text: str, filename: str, urls: Iterable[str]) -> str:
    urls_string = "; ".join(urls) if urls else "No linked documents found"
    return f"""
You are an expert Government Tender Analyst. Extract data from this GeM tender document into a JSON object.

Tender Document: {filename}
---
{text}
---

**Instructions:**
1. Extract all fields. Use "Not Found" if information is missing.
2. For "Go_NoGo_Eligibility_Summary": Provide a clear, action-oriented summary of minimum qualifications (turnover, experience, certifications) required to bid.
3. Standardize financial values to Lakhs/Crores (e.g., "1.8 Crore", "18 Lakhs").
4. For URLs, use: "{urls_string}"

**JSON Schema (use these exact keys):**
{{
  "File_Name": "{filename}",
  "Bid_Number": "e.g., GEM/2025/B/6530170",
  "Item_Category": "Core service/item being procured",
  "Ministry_Department_Office": "Full name of procuring entity",
  "Bid_End_Date": "Date and time bid closes",
  "Bid_Opening_Date": "Date and time technical bid opens",
  "Contract_Period_Months": "Duration in months",
  "Estimated_Value_INR": "Estimated value in Lakhs/Crores",
  "Min_Annual_Turnover_Req_INR": "Required average annual turnover (last 3 years)",
  "Years_Past_Experience_Req": "Years of similar experience required",
  "Required_Certifications": "Mandatory certifications (e.g., CAG Empaneled, ISO)",
  "MSE_Exemption_Available": "Yes/No for MSE exemption",
  "Startup_Exemption_Available": "Yes/No for Startup exemption",
  "EMD_Required_Amount": "EMD amount or Exempted/Not Applicable",
  "Scope_of_Work_Summary": "1-2 sentence summary of work scope",
  "Key_Linked_Documents_URLs": "{urls_string}",
  "Go_NoGo_Eligibility_Summary": "Action-oriented summary of key eligibility criteria"
}}

Return ONLY valid JSON without markdown formatting.
"""


def _build_sow_prompt(text: str, filename: str) -> str:
    return f"""
You are an expert Government Tender Analyst. Extract SOW data from this document into a JSON object.

Document: {filename}
---
{text}
---

**Instructions:**
1. Extract all technical and operational details.
2. Use "Not Found" if information is missing.
3. Provide comprehensive summaries for complex fields.

**JSON Schema (use these exact keys):**
{{
  "File_Name": "{filename}",
  "Bid_Number": "Reference/Bid number if mentioned",
  "Document_Type": "SOW",
  "Technical_Requirements": "All technical specifications",
  "Deliverables": "Expected deliverables list",
  "Timeline_Schedule": "Project timeline and milestones",
  "Performance_Metrics": "KPIs and success criteria",
  "Quality_Standards": "Quality requirements and certifications",
  "Compliance_Requirements": "Regulatory compliance needs",
  "Hardware_Software_Specifications": "Technical specs if applicable",
  "Implementation_Methodology": "Required approach/methodology",
  "Support_and_Maintenance": "Post-implementation support needs",
  "Training_Requirements": "Training for end users/staff",
  "Acceptance_Criteria": "Project acceptance criteria"
}}

Return ONLY valid JSON without markdown formatting.
"""


def call_gemini_json(prompt: str, *, client: genai.Client) -> Dict[str, str] | None:
    """Invoke the Gemini model and coerce the response into JSON."""
    try:
        response = client.models.generate_content(
            model=SELECTED_MODEL,
            contents=prompt,
            generation_config={"response_mime_type": "application/json"},
        )
        return json.loads(response.text)
    except json.JSONDecodeError as exc:
        st.error(f"Failed to parse AI response as JSON: {exc}")
        return None
    except Exception as exc:
        st.error(f"AI Error: {exc}")
        return None


def extract_rfp_data(text: str, filename: str, urls: Iterable[str], *, client: genai.Client):
    prompt = _build_rfp_prompt(text, filename, urls)
    return call_gemini_json(prompt, client=client)


def extract_sow_data(text: str, filename: str, *, client: genai.Client):
    prompt = _build_sow_prompt(text, filename)
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

        progress_bar = st.progress(0)
        status_placeholder = st.empty()

        for index, pdf in enumerate(uploaded_files, start=1):
            status_placeholder.info(f"📄 Processing: **{pdf.name}** ({index}/{len(uploaded_files)})")

            text, urls = extract_text_and_urls_from_pdf(pdf)
            if not text:
                continue

            # Break long documents into manageable chunks for better reliability
            chunks = chunk_text(text)
            combined_text = "\n\n".join(chunks)

            doc_type = document_types.get(pdf.name, "RFP")
            if doc_type == "RFP":
                result = extract_rfp_data(combined_text, pdf.name, urls, client=client)
                if result:
                    rfp_results.append(result)
                    st.success(f"✅ Successfully analyzed {pdf.name} (RFP)")
            else:
                result = extract_sow_data(combined_text, pdf.name, client=client)
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
        json_payload: Dict[str, List[Dict[str, str]]] = {}

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

        st.download_button(
            label="📤 Download Raw JSON",
            data=json.dumps(json_payload, indent=2),
            file_name="rfp_sow_analysis_results.json",
            mime="application/json",
        )

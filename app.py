# app.py - Enhanced RFP Analyzer using Google Gemini

import json
import os
import re
import mimetypes
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple, Union

import pandas as pd
import requests
import streamlit as st
from docx import Document
import fitz  # PyMuPDF
from google import genai

# --- Page Configuration ---
st.set_page_config(
    page_title="RFP Analysis Tool",
    page_icon="📄",
    layout="wide"
)

st.title("📄 RFP Analysis Tool")
st.write(
    "Upload one or more government tender PDFs to automatically extract key information "
    "and manage any supporting Scope of Work (SOW) documents."
)

# --- Constants and Defaults ---
API_KEY_PATH = Path.home() / ".rfp_analyzer_api_key"
DEFAULT_PROMPT = """You are an expert government tender analyst. Read the full RFP carefully and return a single JSON object only.
Capture the following baseline fields as strings (join lists with \n where needed).
- bid_number
- start_date
- end_date
- city
- organisation_name
- department_name
- period_of_contract
- estimated_bid_amount
- category_or_type
- scope_of_work (array of bullet points that retain the concrete activities, deliverables, frequencies, timelines, and compliance responsibilities. Keep one bullet per obligation and preserve numbering or lettering.)
- eligibility_criteria (array of bullet points that spell out turnover values, certifications, manpower mix, past experience, geography, registrations, or any disqualifications verbatim. Include thresholds even if repeated elsewhere.)
- delivery_or_execution_mode (capture onsite/offsite/hybrid requirements and any visit frequency expectations)
- contact_details
- minimum_turnover_requirement
- experience_requirement
- earnest_money_deposit
- bid_security
- important_dates (array; include bid opening, pre-bid, clarification deadlines, contract award, etc.)

Additionally:
- Place every other commercial, technical, submission, penalty, payment, staffing, document checklist, or evaluation requirement inside an "additional_fields" object using descriptive keys.
- Do not omit financial numbers such as turnover thresholds, bid values, security deposits, or penalties even if they appear in multiple sections.
- Preserve numbering and sub-bullets exactly as written within list items.
- Respond with valid JSON only. Do not include markdown, commentary, or triple backticks.
- If a field is absent, set it to "Not Found" (or an empty array for list fields)."""
DOCUMENT_FILE_EXTENSIONS = {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".zip"}
DOCUMENT_LINK_KEYWORDS: Tuple[str, ...] = (
    "sow",
    "scope",
    "atc",
    "additional",
    "annex",
    "corrigendum",
    "tender",
    "document",
    "attachment",
)

if "analysis_results" not in st.session_state:
    st.session_state["analysis_results"] = []
if "attachment_cache" not in st.session_state:
    st.session_state["attachment_cache"] = {}
if "sow_links_cache" not in st.session_state:
    st.session_state["sow_links_cache"] = {}
if "fetched_documents" not in st.session_state:
    st.session_state["fetched_documents"] = {}
if "analysis_notifications" not in st.session_state:
    st.session_state["analysis_notifications"] = []

# --- Utility Functions ---
def load_api_key() -> Optional[str]:
    """Load the Gemini API key from environment, local storage, or session."""
    env_key = os.getenv("GENAI_API_KEY")
    if env_key:
        return env_key.strip()

    if API_KEY_PATH.exists():
        try:
            return API_KEY_PATH.read_text(encoding="utf-8").strip()
        except OSError:
            st.warning("Found a stored API key but could not read it. Please re-enter below.")

    return st.session_state.get("api_key")


def save_api_key(key: str) -> None:
    """Persist the API key so it only needs to be entered once."""
    key = key.strip()
    if not key:
        st.error("API key cannot be empty.")
        return

    API_KEY_PATH.write_text(key, encoding="utf-8")
    try:
        os.chmod(API_KEY_PATH, 0o600)
    except OSError:
        # On some systems (e.g. Windows) chmod may fail; ignore silently.
        pass
    st.session_state["api_key"] = key
    st.success("API key saved. It will be reused automatically next time.")


def get_client(api_key: str) -> genai.Client:
    return genai.Client(api_key=api_key)


def extract_text_from_pdf(pdf_bytes: bytes) -> Optional[str]:
    try:
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        text = "\n".join(page.get_text() for page in pdf_document)
        pdf_document.close()
        return text
    except Exception as exc:
        st.error(f"Error reading PDF: {exc}")
        return None


def extract_text_from_docx(docx_bytes: bytes) -> Optional[str]:
    try:
        document = Document(BytesIO(docx_bytes))
        return "\n".join(paragraph.text for paragraph in document.paragraphs)
    except Exception as exc:
        st.error(f"Unable to read DOCX file: {exc}")
        return None


def extract_attachments_from_pdf(pdf_bytes: bytes) -> List[Tuple[str, bytes, Optional[str]]]:
    """Return embedded files inside the PDF (name, data, description)."""
    attachments: List[Tuple[str, bytes, Optional[str]]] = []
    try:
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        getter = getattr(pdf_document, "embeddedFileGet", None)
        if callable(getter):
            count = getattr(pdf_document, "embeddedFileCount", 0) or 0
            names = list(getattr(pdf_document, "embeddedFileNames", lambda: [])())
            for index in range(count):
                info = None
                try:
                    info = getter(index)
                except TypeError:
                    if index < len(names):
                        info = getter(names[index])
                if not isinstance(info, dict):
                    continue
                name = info.get("name") or (names[index] if index < len(names) else f"attachment_{index + 1}")
                file_bytes = info.get("file", b"")
                description = info.get("desc")
                if file_bytes:
                    attachments.append((name, file_bytes, description))
        elif hasattr(pdf_document, "embeddedFileNames"):
            for name in pdf_document.embeddedFileNames():
                info = pdf_document.embeddedFileGet(name)
                file_bytes = info.get("file", b"") if isinstance(info, dict) else b""
                if not file_bytes:
                    continue
                attachments.append((info.get("name", name), file_bytes, info.get("desc") if isinstance(info, dict) else None))
        pdf_document.close()
    except Exception as exc:
        st.warning(f"Unable to extract embedded files: {exc}")
    return attachments


def find_reference_links(text: str) -> List[str]:
    url_pattern = re.compile(r"https?://[^\s>]+", re.IGNORECASE)
    links = []
    for match in url_pattern.findall(text):
        cleaned = match.rstrip(').,;"\'\n')
        links.append(cleaned)
    return list(dict.fromkeys(links))  # remove duplicates while preserving order


def _link_matches_keywords(url: str, keywords: Iterable[str]) -> bool:
    lowered = url.lower()
    return any(keyword in lowered for keyword in keywords)


def fetch_external_document(url: str) -> Optional[Tuple[str, bytes, Optional[str]]]:
    try:
        response = requests.get(url, timeout=20)
        response.raise_for_status()
        filename = url.split("/")[-1] or "downloaded_file"
        if "content-disposition" in response.headers:
            disposition = response.headers["content-disposition"]
            match = re.search(r'filename="?([^";]+)"?', disposition)
            if match:
                filename = match.group(1)
        content_type = response.headers.get("content-type")
        return filename, response.content, content_type
    except requests.RequestException as exc:
        st.error(f"Failed to download {url}: {exc}")
        return None


def run_model(prompt_template: str, text_content: str, api_key: str) -> str:
    client = get_client(api_key)
    prompt = f"{prompt_template.strip()}\n\n---\n{text_content}\n---"
    response = client.models.generate_content(
        model="models/gemini-2.5-pro",
        contents=prompt
    )
    return response.text


def _ensure_json_payload(raw_text: str) -> Optional[Union[Dict, List]]:
    """Attempt to extract and decode JSON from the model output."""
    raw_text = raw_text.strip()
    if not raw_text:
        return None

    json_snippet = raw_text
    if not raw_text.lstrip().startswith(("{", "[")):
        try:
            start = raw_text.index("{")
            end = raw_text.rindex("}") + 1
            json_snippet = raw_text[start:end]
        except ValueError:
            return None

    try:
        return json.loads(json_snippet)
    except json.JSONDecodeError:
        return None


def _humanize_key(key: str) -> str:
    key = key.replace("_", " ").replace("-", " ")
    return key.strip().title()


def _normalize_text(value: Union[str, int, float]) -> str:
    text = str(value)
    text = text.replace("**", "")
    text = text.replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _list_to_lines(items: List[Any], level: int = 0) -> List[str]:
    lines: List[str] = []
    for item in items:
        if item is None:
            continue
        if isinstance(item, list):
            lines.extend(_list_to_lines(item, level + 1))
            continue
        if isinstance(item, dict):
            dict_parts: List[str] = []
            for key, value in item.items():
                inner_text = clean_text_value(value)
                if inner_text:
                    dict_parts.append(f"{_humanize_key(str(key))}: {inner_text}")
            if dict_parts:
                prefix = "• " if level == 0 else "  " * level + "- "
                lines.append(f"{prefix}{'; '.join(dict_parts)}")
            continue
        text = _normalize_text(item)
        if not text:
            continue
        sublines = [part.strip() for part in text.split("\n") if part.strip()]
        if not sublines:
            continue
        first_prefix = "• " if level == 0 else "  " * level + "- "
        continuation_prefix = "  - " if level == 0 else "  " * (level + 1) + "- "
        lines.append(f"{first_prefix}{sublines[0]}")
        for extra in sublines[1:]:
            lines.append(f"{continuation_prefix}{extra}")
    return lines


def clean_text_value(value: Union[str, List, Dict[str, str], None]) -> str:
    if value is None:
        return ""
    if isinstance(value, list):
        bullet_lines = _list_to_lines(value)
        return "\n".join(bullet_lines)
    if isinstance(value, dict):
        parts = [f"{_humanize_key(str(k))}: {clean_text_value(v)}" for k, v in value.items()]
        return "\n".join(part for part in parts if part.strip())
    return _normalize_text(value)


BASE_FIELD_LABELS: Dict[str, str] = {
    "bid_number": "Bid Number",
    "start_date": "Start Date",
    "end_date": "End Date",
    "city": "City",
    "organisation_name": "Organization",
    "organization_name": "Organization",
    "period_of_contract": "Period Of Contract",
    "estimated_bid_amount": "Est. Bid Amount (Rs.)",
    "category_or_type": "Service Category",
    "scope_of_work": "Scope Of Work",
    "scope": "Scope Of Work",
    "eligibility_criteria": "Eligibility Criteria",
    "delivery_or_execution_mode": "Offsite / Onsite",
    "minimum_turnover_requirement": "Minimum Turnover Requirement",
    "experience_requirement": "Experience Requirement",
    "earnest_money_deposit": "Earnest Money Deposit",
    "bid_security": "Bid Security",
    "important_dates": "Important Dates",
    "contact_details": "Contact Details",
}

def parse_ai_response(response_text: str) -> Dict[str, str]:
    data: Dict[str, str] = {}
    payload = _ensure_json_payload(response_text)
    if isinstance(payload, dict):
        additional_entries: List[str] = []
        additional_fields = payload.pop("additional_fields", {})
        for key, value in payload.items():
            label = BASE_FIELD_LABELS.get(key, _humanize_key(key))
            text_value = clean_text_value(value)
            if key in BASE_FIELD_LABELS:
                data[label] = text_value
            else:
                if text_value:
                    additional_entries.append(f"{label}: {text_value}")
        if isinstance(additional_fields, dict):
            for key, value in additional_fields.items():
                text_value = clean_text_value(value)
                if text_value:
                    additional_entries.append(f"{_humanize_key(key)}: {text_value}")
        if additional_entries:
            data["Additional Details"] = "\n".join(additional_entries)
        for label in BASE_FIELD_LABELS.values():
            data.setdefault(label, "Not Found")
    else:
        lines = response_text.strip().split("\n")
        for line in lines:
            if ":" in line:
                key, value = line.split(":", 1)
                data[key.strip()] = clean_text_value(value)
        if response_text and "Observations" not in data:
            data["Observations"] = clean_text_value(response_text)
        for label in BASE_FIELD_LABELS.values():
            data.setdefault(label, "Not Found")
    return data


def add_result_to_state(result: Dict[str, str]) -> None:
    results: List[Dict[str, str]] = st.session_state.setdefault("analysis_results", [])
    identifier = (result.get("File Name"), result.get("Bid Number"), result.get("Source Document"))
    for existing in results:
        existing_id = (
            existing.get("File Name"),
            existing.get("Bid Number"),
            existing.get("Source Document"),
        )
        if existing_id == identifier:
            existing.update(result)
            break
    else:
        results.append(result)


def _append_notification(level: str, message: str) -> None:
    st.session_state.setdefault("analysis_notifications", []).append((level, message))


def analyze_and_store(
    name: str,
    data: bytes,
    prompt_template: str,
    source_label: Optional[str],
) -> None:
    active_api_key = st.session_state.get("api_key") or load_api_key()
    if not active_api_key:
        _append_notification("warning", "Please provide your Gemini API key before analyzing supporting documents.")
        return

    result = analyze_document(
        name,
        data,
        prompt_template,
        active_api_key,
        source_label=source_label,
    )
    if result:
        add_result_to_state(result)
        _append_notification("success", f"Added structured data from {name}.")


def analyze_document(
    name: str,
    data: bytes,
    prompt_template: str,
    api_key: str,
    source_label: Optional[str] = None,
    text_override: Optional[str] = None,
) -> Optional[Dict[str, str]]:
    suffix = Path(name).suffix.lower()
    text: Optional[str] = text_override

    if text is None:
        if suffix == ".pdf":
            text = extract_text_from_pdf(data)
        elif suffix in {".doc", ".docx"}:
            text = extract_text_from_docx(data)
        else:
            try:
                text = data.decode("utf-8")
            except UnicodeDecodeError:
                st.warning(f"Cannot extract text from {name}. Unsupported format.")
                return None

    if not text:
        return None

    try:
        response_text = run_model(prompt_template, text, api_key)
    except Exception as exc:
        st.error(f"AI analysis failed for {name}: {exc}")
        return None

    parsed_data = parse_ai_response(response_text)
    parsed_data['File Name'] = name
    if source_label:
        parsed_data['Source Document'] = source_label
    return parsed_data


# --- Sidebar Configuration ---
st.sidebar.header("Configuration")

stored_key = load_api_key()
if stored_key:
    st.sidebar.success("API key loaded from secure storage.")
    if "api_key" not in st.session_state:
        st.session_state["api_key"] = stored_key
else:
    st.sidebar.info("Enter your Gemini API key once. It will be stored securely on this machine.")

with st.sidebar.expander("Manage API Key", expanded=not bool(stored_key)):
    api_key_input = st.text_input(
        "Google Gemini API Key",
        value=stored_key or "",
        type="password"
    )
    if st.button("Save API Key"):
        save_api_key(api_key_input)
        stored_key = api_key_input.strip()

prompt_template = st.sidebar.text_area(
    "Prompt Template",
    value=st.session_state.get("prompt_template", DEFAULT_PROMPT),
    height=240,
)

if st.sidebar.button("Reset Prompt to Default"):
    prompt_template = DEFAULT_PROMPT

st.session_state["prompt_template"] = prompt_template

st.sidebar.caption(
    "Tweak the prompt template to try different extraction instructions."
)

# --- File Upload ---
uploaded_files = st.file_uploader(
    "Choose RFP PDF files",
    type=["pdf"],
    accept_multiple_files=True
)

st.write("\n")

# --- Analysis Trigger ---
analyze_clicked = st.button("🚀 Analyze RFPs", type="primary")

if analyze_clicked:
    active_api_key = stored_key or st.session_state.get("api_key")
    if not active_api_key:
        st.warning("Please provide your Gemini API key in the sidebar before analyzing.")
    elif not uploaded_files:
        st.warning("Please upload at least one PDF file.")
    else:
        st.session_state["analysis_results"] = []
        st.session_state["fetched_documents"] = {}
        attachments_cache: Dict[str, List[Tuple[str, bytes, Optional[str]]]] = {}
        sow_links_cache: Dict[str, List[str]] = {}
        with st.spinner("Analyzing documents... Please wait."):
            for pdf_file in uploaded_files:
                st.write("---")
                st.subheader(f"📑 {pdf_file.name}")
                pdf_bytes = pdf_file.getvalue()
                text = extract_text_from_pdf(pdf_bytes)
                if not text:
                    st.warning(f"Could not extract text from {pdf_file.name}.")
                    continue

                result = analyze_document(
                    pdf_file.name,
                    pdf_bytes,
                    prompt_template,
                    active_api_key,
                    source_label="Primary Upload",
                    text_override=text,
                )
                if result:
                    add_result_to_state(result)
                    st.success(f"Captured structured data from {pdf_file.name}.")

                attachments = extract_attachments_from_pdf(pdf_bytes)
                if attachments:
                    attachments_cache[pdf_file.name] = attachments

                links = find_reference_links(text)
                sow_links: List[str] = []
                for link in links:
                    if (
                        Path(link.lower()).suffix in DOCUMENT_FILE_EXTENSIONS
                        or _link_matches_keywords(link, DOCUMENT_LINK_KEYWORDS)
                    ):
                        sow_links.append(link)
                        cache_key = f"{pdf_file.name}||{link}"
                        if cache_key not in st.session_state["fetched_documents"]:
                            fetched = fetch_external_document(link)
                            if fetched:
                                fetched_name, fetched_bytes, content_type = fetched
                                st.session_state["fetched_documents"][cache_key] = {
                                    "name": fetched_name,
                                    "bytes": fetched_bytes,
                                    "mime": content_type or mimetypes.guess_type(fetched_name)[0],
                                }
                                _append_notification(
                                    "info",
                                    f"Fetched supporting document '{fetched_name}'. Use the download button to save or analyse it.",
                                )
                if sow_links:
                    sow_links_cache[pdf_file.name] = sow_links

        st.session_state["attachment_cache"] = attachments_cache
        st.session_state["sow_links_cache"] = sow_links_cache

        if st.session_state["analysis_results"]:
            st.success("Analysis complete. Review the sections below for structured results, attachments, and downloads.")
        else:
            st.info("Analysis finished but no structured data was returned. Consider refining the prompt and trying again.")

attachment_cache = st.session_state.get("attachment_cache", {})
if attachment_cache:
    st.write("---")
    st.header("Embedded Attachments")
    for base_name, attachments in attachment_cache.items():
        with st.expander(f"📎 Attachments in {base_name}"):
            for idx, (att_name, att_data, desc) in enumerate(attachments, start=1):
                st.write(f"{idx}. {att_name}")
                if desc:
                    st.caption(desc)
                mime_type = mimetypes.guess_type(att_name)[0] or "application/octet-stream"
                st.download_button(
                    label="Download & Analyze",
                    data=att_data,
                    file_name=att_name,
                    mime=mime_type,
                    key=f"download_{base_name}_{idx}",
                    on_click=analyze_and_store,
                    kwargs={
                        "name": att_name,
                        "data": att_data,
                        "prompt_template": prompt_template,
                        "source_label": f"Attachment from {base_name}",
                    },
                )

sow_cache = st.session_state.get("sow_links_cache", {})
if sow_cache:
    st.write("---")
    st.header("Referenced External Documents")
    for base_name, links in sow_cache.items():
        with st.expander(f"🔗 Links mentioned in {base_name}"):
            for idx, link in enumerate(links, start=1):
                st.markdown(f"{idx}. [{link}]({link})")
                cache_key = f"{base_name}||{link}"
                stored_file = st.session_state["fetched_documents"].get(cache_key)
                if stored_file:
                    st.download_button(
                        label=f"Download & Analyze {stored_file['name']}",
                        data=stored_file["bytes"],
                        file_name=stored_file["name"],
                        mime=stored_file.get("mime") or "application/octet-stream",
                        key=f"download_link_{idx}_{base_name}",
                        on_click=analyze_and_store,
                        kwargs={
                            "name": stored_file["name"],
                            "data": stored_file["bytes"],
                            "prompt_template": prompt_template,
                            "source_label": f"Link from {base_name}",
                        },
                    )
                else:
                    fetch_key = f"fetch_{base_name}_{idx}"
                    if st.button("Fetch Document", key=fetch_key):
                        fetched = fetch_external_document(link)
                        if fetched:
                            fetched_name, fetched_bytes, content_type = fetched
                            st.session_state["fetched_documents"][cache_key] = {
                                "name": fetched_name,
                                "bytes": fetched_bytes,
                                "mime": content_type or mimetypes.guess_type(fetched_name)[0],
                            }
                            _append_notification(
                                "success",
                                f"Fetched {fetched_name}. Use the download button to save or analyse it.",
                            )

notifications = st.session_state.get("analysis_notifications", [])
if notifications:
    for level, message in notifications:
        if level == "success":
            st.success(message)
        elif level == "warning":
            st.warning(message)
        elif level == "error":
            st.error(message)
        else:
            st.info(message)
    st.session_state["analysis_notifications"] = []

analysis_results = st.session_state.get("analysis_results", [])
if analysis_results:
    st.write("---")
    st.header("Structured Analysis Results")
    for record in analysis_results:
        title = record.get("File Name") or record.get("Source Document") or "Document"
        with st.expander(f"📑 {title}", expanded=False):
            st.json(record)

    df = pd.DataFrame(analysis_results).fillna("")
    preferred_order = [
        "File Name",
        "Source Document",
        "Bid Number",
        "Start Date",
        "End Date",
        "City",
        "Organization",
        "Period Of Contract",
        "Est. Bid Amount (Rs.)",
        "Service Category",
        "Scope Of Work",
        "Eligibility Criteria",
        "Offsite / Onsite",
        "Minimum Turnover Requirement",
        "Experience Requirement",
        "Earnest Money Deposit",
        "Bid Security",
        "Important Dates",
        "Contact Details",
        "Additional Details",
    ]
    for col in preferred_order:
        if col not in df.columns:
            df[col] = ""

    ordered_columns = []
    for col in preferred_order:
        if col in df.columns and col not in ordered_columns:
            ordered_columns.append(col)
    for col in df.columns:
        if col not in ordered_columns:
            ordered_columns.append(col)
    df = df[ordered_columns]

    st.dataframe(df, use_container_width=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='RFP_Data')

    excel_data = output.getvalue()

    st.download_button(
        label="📥 Download Data as Excel",
        data=excel_data,
        file_name="rfp_analysis_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Upload an RFP and run the analysis to see results here.")


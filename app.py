# app.py - Enhanced RFP Analyzer with Two-Step Process

import json
import os
import re
import mimetypes
import zipfile
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

st.title("📄 RFP Analysis Automation Tool")
st.write(
    "Upload government tender PDFs to extract key information in two steps: "
    "first analyze the main RFP, then optionally analyze supporting documents."
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
    "sow", "scope", "atc", "additional", "annex", "corrigendum", "tender", "document", "attachment",
)

# Initialize session state
if "primary_analysis_results" not in st.session_state:
    st.session_state["primary_analysis_results"] = []
if "supporting_analysis_results" not in st.session_state:
    st.session_state["supporting_analysis_results"] = []
if "extracted_links" not in st.session_state:
    st.session_state["extracted_links"] = {}
if "attachment_cache" not in st.session_state:
    st.session_state["attachment_cache"] = {}
if "current_step" not in st.session_state:
    st.session_state["current_step"] = 1  # 1 or 2

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

def extract_embedded_links_from_pdf(pdf_bytes: bytes) -> List[Tuple[str, str, str, str]]:
    """Extract embedded hyperlinks from PDF annotations (url, link_text, page_number, label)."""
    embedded_links = []
    seen = set()
    
    try:
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        for page_num, page in enumerate(pdf_document, start=1):
            page_text = page.get_text()
            links = page.get_links()
            
            for link in links:
                if link.get("kind") == 2:  # URI link
                    uri = link.get("uri", "")
                    if uri and uri not in seen:
                        seen.add(uri)
                        rect = link.get("from")
                        link_text = "Click here"
                        label = "Document"
                        
                        if rect:
                            try:
                                text_instances = page.get_text("words")
                                nearby_words = []
                                
                                for word in text_instances:
                                    word_rect = fitz.Rect(word[:4])
                                    if word_rect.intersects(rect) or word_rect.x0 <= rect.x1 + 100:
                                        nearby_words.append(word[4])
                                
                                if nearby_words:
                                    context = " ".join(nearby_words[:15])
                                    context_lower = context.lower()
                                    
                                    if "atc" in context_lower:
                                        label = "ATC Document"
                                    elif "sow" in context_lower or "scope of work" in context_lower:
                                        label = "Scope of Work (SOW)"
                                    elif "corrigendum" in context_lower:
                                        label = "Corrigendum"
                                    elif "annex" in context_lower or "annexure" in context_lower:
                                        label = "Annexure"
                                    elif "tender" in context_lower:
                                        label = "Tender Document"
                                    
                                    link_text = context[:60]
                            except:
                                pass
                        
                        embedded_links.append((uri, link_text, f"Page {page_num}", label))
        pdf_document.close()
    except Exception as exc:
        st.warning(f"Unable to extract embedded links: {exc}")
    
    return embedded_links

def find_reference_links(text: str) -> List[Tuple[str, str]]:
    """Extract plain text URLs and categorize them."""
    url_pattern = re.compile(r"https?://[^\s>]+", re.IGNORECASE)
    links = []
    seen = set()
    
    for match in url_pattern.findall(text):
        cleaned = match.rstrip(').,;"\'\n')
        if cleaned in seen:
            continue
        seen.add(cleaned)
        
        lowered = cleaned.lower()
        if any(keyword in lowered for keyword in DOCUMENT_LINK_KEYWORDS):
            category = "Related Document"
        elif Path(lowered).suffix in DOCUMENT_FILE_EXTENSIONS:
            category = "Other Document"
        else:
            category = "General Link"
        
        links.append((cleaned, category))
    
    return links

def categorize_link(url: str, context: str = "", label: str = "") -> str:
    """Categorize a URL based on its content, context, and label."""
    lowered = url.lower()
    context_lower = context.lower()
    label_lower = label.lower()
    
    if "atc" in label_lower:
        return "ATC Document"
    if "sow" in label_lower or "scope of work" in label_lower:
        return "SOW Document"
    if "corrigendum" in label_lower:
        return "Corrigendum"
    if "annex" in label_lower:
        return "Annexure"
    if "atc" in lowered or "atc" in context_lower:
        return "ATC Document"
    if "sow" in lowered or "scope" in lowered or "sow" in context_lower or "scope of work" in context_lower:
        return "SOW Document"
    if "corrigendum" in lowered or "corrigendum" in context_lower:
        return "Corrigendum"
    if any(keyword in lowered for keyword in DOCUMENT_LINK_KEYWORDS):
        return "Related Document"
    elif Path(lowered).suffix in DOCUMENT_FILE_EXTENSIONS:
        return "Other Document"
    else:
        return "General Link"

def fetch_external_document(url: str) -> Optional[Tuple[str, bytes, Optional[str]]]:
    """Download document from URL."""
    try:
        response = requests.get(url, timeout=30)
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
        model="gemini-2.0-flash-exp",
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

def normalize_delivery_mode(raw_text: str) -> Tuple[str, Optional[str]]:
    """Collapse delivery/execution notes to Onsite/Offsite/Hybrid labels."""
    if not raw_text:
        return "Not Found", None
    cleaned = raw_text.strip()
    if cleaned.lower() in {"not found", "na", "n/a"}:
        return "Not Found", None
    lowered = cleaned.lower()
    onsite = bool(re.search(r"\bon[-\s]?site\b", lowered))
    offsite = bool(
        re.search(r"\boff[-\s]?site\b", lowered)
        or "remote" in lowered
        or "work from home" in lowered
    )
    if onsite and offsite:
        return "Hybrid", cleaned
    if onsite:
        return "Onsite", cleaned if len(cleaned) > len("Onsite") else None
    if offsite:
        return "Offsite", cleaned if len(cleaned) > len("Offsite") else None
    return "Not Found", cleaned

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
                if key == "delivery_or_execution_mode":
                    normalized, detail = normalize_delivery_mode(text_value)
                    data[label] = normalized
                    if detail and detail != normalized:
                        additional_entries.append(f"Delivery Details: {detail}")
                else:
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

def create_excel_output(primary_results: List[Dict], supporting_results: List[Dict] = None) -> bytes:
    """Create structured Excel output with multiple sheets."""
    output = BytesIO()
    
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Primary RFPs sheet
        if primary_results:
            primary_df = pd.DataFrame(primary_results).fillna("")
            primary_df = order_dataframe_columns(primary_df)
            primary_df.to_excel(writer, index=False, sheet_name='Primary RFPs')
            
            # Format the sheet
            workbook = writer.book
            sheet = writer.sheets['Primary RFPs']
            format_worksheet(sheet)
        
        # Supporting documents sheet
        if supporting_results:
            supporting_df = pd.DataFrame(supporting_results).fillna("")
            supporting_df = order_dataframe_columns(supporting_df)
            supporting_df.to_excel(writer, index=False, sheet_name='Supporting Docs')
            
            sheet = writer.sheets['Supporting Docs']
            format_worksheet(sheet)
        
        # Combined summary if both exist
        if primary_results and supporting_results:
            combined_df = pd.concat([
                pd.DataFrame(primary_results),
                pd.DataFrame(supporting_results)
            ], ignore_index=True).fillna("")
            combined_df = order_dataframe_columns(combined_df)
            combined_df.to_excel(writer, index=False, sheet_name='All Documents')
            
            sheet = writer.sheets['All Documents']
            format_worksheet(sheet)
        
        # Metadata sheet
        total_docs = len(primary_results) + (len(supporting_results) if supporting_results else 0)
        metadata_df = pd.DataFrame({
            'Metric': [
                'Total Documents Analyzed',
                'Primary RFPs',
                'Supporting Documents',
                'Analysis Date'
            ],
            'Value': [
                total_docs,
                len(primary_results),
                len(supporting_results) if supporting_results else 0,
                pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
            ]
        })
        metadata_df.to_excel(writer, index=False, sheet_name='Metadata')
    
    return output.getvalue()

def order_dataframe_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Order DataFrame columns in a structured way."""
    column_structure = {
        "Document Info": ["File Name", "Source Document", "Bid Number"],
        "Timeline": ["Start Date", "End Date", "Period Of Contract", "Important Dates"],
        "Organization": ["Organization", "City", "Contact Details"],
        "Financial": [
            "Est. Bid Amount (Rs.)",
            "Minimum Turnover Requirement",
            "Earnest Money Deposit",
            "Bid Security"
        ],
        "Service Details": [
            "Service Category",
            "Scope Of Work",
            "Offsite / Onsite",
            "Experience Requirement"
        ],
        "Eligibility": ["Eligibility Criteria"],
        "Additional": ["Additional Details"]
    }
    
    ordered_columns = []
    for group_cols in column_structure.values():
        for col in group_cols:
            if col in df.columns and col not in ordered_columns:
                ordered_columns.append(col)
    
    for col in df.columns:
        if col not in ordered_columns:
            ordered_columns.append(col)
    
    for col in ordered_columns:
        if col not in df.columns:
            df[col] = "Not Found"
    
    return df[ordered_columns]

def format_worksheet(sheet):
    """Format worksheet with proper column widths."""
    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        sheet.column_dimensions[column_letter].width = adjusted_width

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

# --- STEP 1: Analyze Primary RFP Documents ---
st.header("Step 1: Analyze Primary RFP Documents")
st.write("Upload your main RFP PDF files and analyze them to get structured data.")

uploaded_files = st.file_uploader(
    "Choose RFP PDF files",
    type=["pdf"],
    accept_multiple_files=True,
    key="primary_upload"
)

if st.button("Analyze Primary RFPs", type="primary", disabled=not uploaded_files):
    active_api_key = stored_key or st.session_state.get("api_key")
    if not active_api_key:
        st.warning("Please provide your Gemini API key in the sidebar.")
    else:
        st.session_state["primary_analysis_results"] = []
        st.session_state["extracted_links"] = {}
        st.session_state["attachment_cache"] = {}
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, pdf_file in enumerate(uploaded_files):
            status_text.text(f"Analyzing {pdf_file.name}...")
            
            pdf_bytes = pdf_file.getvalue()
            text = extract_text_from_pdf(pdf_bytes)
            
            if not text:
                st.warning(f"Could not extract text from {pdf_file.name}.")
                continue
            
            # Analyze the main RFP
            result = analyze_document(
                pdf_file.name,
                pdf_bytes,
                prompt_template,
                active_api_key,
                source_label="Primary RFP",
                text_override=text,
            )
            
            if result:
                st.session_state["primary_analysis_results"].append(result)
            
            # Extract links and attachments for Step 2
            embedded_links = extract_embedded_links_from_pdf(pdf_bytes)
            text_links = find_reference_links(text)
            
            all_links = []
            for url, link_text, page_info, label in embedded_links:
                category = categorize_link(url, link_text, label)
                all_links.append((url, category, f"{link_text} ({page_info})", label))
            
            for url, category in text_links:
                if not any(link[0] == url for link in all_links):
                    all_links.append((url, category, "Plain text URL", ""))
            
            if all_links:
                st.session_state["extracted_links"][pdf_file.name] = all_links
            
            # Extract attachments
            attachments = extract_attachments_from_pdf(pdf_bytes)
            if attachments:
                st.session_state["attachment_cache"][pdf_file.name] = attachments
            
            progress_bar.progress((idx + 1) / len(uploaded_files))
        
        status_text.text("Analysis complete!")
        st.session_state["current_step"] = 1
        st.success(f"Analyzed {len(st.session_state['primary_analysis_results'])} primary RFP documents.")
        st.rerun()

# Display Step 1 Results
if st.session_state["primary_analysis_results"]:
    st.write("---")
    st.subheader("Primary RFP Analysis Results")
    
    primary_df = pd.DataFrame(st.session_state["primary_analysis_results"]).fillna("")
    primary_df = order_dataframe_columns(primary_df)
    
    st.dataframe(primary_df, use_container_width=True)
    
    # Download primary results
    excel_data = create_excel_output(st.session_state["primary_analysis_results"])
    
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="Download Primary RFP Analysis (Excel)",
            data=excel_data,
            file_name=f"primary_rfp_analysis_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    
    with col2:
        csv_data = primary_df.to_csv(index=False)
        st.download_button(
            label="Download Primary RFP Analysis (CSV)",
            data=csv_data,
            file_name=f"primary_rfp_analysis_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
        )
    
    # Show summary of available supporting documents
    total_links = sum(len(links) for links in st.session_state.get("extracted_links", {}).values())
    total_attachments = sum(len(atts) for atts in st.session_state.get("attachment_cache", {}).values())
    
    if total_links > 0 or total_attachments > 0:
        st.info(f"Found {total_links} document links and {total_attachments} embedded attachments. Proceed to Step 2 to analyze them.")

# --- STEP 2: Analyze Supporting Documents ---
if st.session_state["primary_analysis_results"]:
    st.write("---")
    st.header("Step 2: Analyze Supporting Documents (Optional)")
    st.write("Download and analyze ATC, SOW, and other supporting documents referenced in your RFPs.")
    
    # Section 2A: Embedded Attachments
    attachment_cache = st.session_state.get("attachment_cache", {})
    if attachment_cache:
        st.subheader("2A: Embedded Attachments")
        st.write("These files are embedded directly in the PDF:")
        
        for pdf_name, attachments in attachment_cache.items():
            with st.expander(f"Attachments in {pdf_name}", expanded=True):
                for idx, (att_name, att_data, desc) in enumerate(attachments, start=1):
                    col1, col2, col3 = st.columns([2, 2, 2])
                    
                    with col1:
                        st.write(f"**{att_name}**")
                        if desc:
                            st.caption(desc)
                    
                    with col2:
                        file_size = len(att_data) / 1024
                        st.caption(f"Size: {file_size:.1f} KB")
                    
                    with col3:
                        mime_type = mimetypes.guess_type(att_name)[0] or "application/octet-stream"
                        
                        col3a, col3b = st.columns(2)
                        with col3a:
                            st.download_button(
                                label="Download",
                                data=att_data,
                                file_name=att_name,
                                mime=mime_type,
                                key=f"dl_att_{pdf_name}_{idx}",
                            )
                        
                        with col3b:
                            if st.button("Analyze", key=f"analyze_att_{pdf_name}_{idx}"):
                                active_api_key = stored_key or st.session_state.get("api_key")
                                if active_api_key:
                                    with st.spinner(f"Analyzing {att_name}..."):
                                        result = analyze_document(
                                            att_name,
                                            att_data,
                                            prompt_template,
                                            active_api_key,
                                            source_label=f"Attachment from {pdf_name}",
                                        )
                                        if result:
                                            st.session_state["supporting_analysis_results"].append(result)
                                            st.success(f"Analyzed {att_name}")
                                            st.rerun()
    
    # Section 2B: External Links
    extracted_links = st.session_state.get("extracted_links", {})
    if extracted_links:
        st.write("---")
        st.subheader("2B: External Document Links")
        st.write("Select documents to download and analyze:")
        
        # Temporary storage for selected links and downloaded files
        if "temp_selected_links" not in st.session_state:
            st.session_state["temp_selected_links"] = {}
        if "downloaded_files" not in st.session_state:
            st.session_state["downloaded_files"] = {}
        
        for pdf_name, links in extracted_links.items():
            with st.expander(f"Links from {pdf_name}", expanded=True):
                if pdf_name not in st.session_state["temp_selected_links"]:
                    st.session_state["temp_selected_links"][pdf_name] = []
                
                for idx, link_data in enumerate(links):
                    if len(link_data) >= 4:
                        url, category, context, label = link_data[:4]
                    elif len(link_data) == 3:
                        url, category, context = link_data
                        label = ""
                    else:
                        url, category = link_data
                        context = ""
                        label = ""
                    
                    col1, col2, col3, col4 = st.columns([0.5, 3, 1.5, 1])
                    
                    with col1:
                        is_selected = st.checkbox(
                            "Select",
                            key=f"select_link_{pdf_name}_{idx}",
                            value=url in st.session_state["temp_selected_links"][pdf_name]
                        )
                        
                        if is_selected and url not in st.session_state["temp_selected_links"][pdf_name]:
                            st.session_state["temp_selected_links"][pdf_name].append(url)
                        elif not is_selected and url in st.session_state["temp_selected_links"][pdf_name]:
                            st.session_state["temp_selected_links"][pdf_name].remove(url)
                    
                    with col2:
                        if label and label != "Document":
                            st.markdown(f"**{label}**")
                        st.markdown(f"[{url}]({url})")
                        if context:
                            st.caption(context)
                    
                    with col3:
                        if category == "ATC Document":
                            st.markdown("ATC")
                        elif category == "SOW Document":
                            st.markdown("SOW")
                        elif category == "Corrigendum":
                            st.markdown("Corrigendum")
                        elif category == "Annexure":
                            st.markdown("Annexure")
                        elif category == "Related Document":
                            st.markdown("Related")
                        else:
                            st.markdown(category)
                    
                    with col4:
                        # Show download button if file has been fetched
                        file_key = f"{pdf_name}||{url}"
                        if file_key in st.session_state["downloaded_files"]:
                            file_info = st.session_state["downloaded_files"][file_key]
                            st.download_button(
                                label="Download",
                                data=file_info["bytes"],
                                file_name=file_info["name"],
                                mime=file_info.get("mime", "application/octet-stream"),
                                key=f"dl_ext_{pdf_name}_{idx}",
                            )
        
        # Analyze selected button
        st.write("---")
        total_selected = sum(len(links) for links in st.session_state["temp_selected_links"].values())
        
        if st.button(f"Download & Analyze Selected Documents ({total_selected})", type="primary", disabled=total_selected == 0):
            active_api_key = stored_key or st.session_state.get("api_key")
            
            if total_selected == 0:
                st.warning("Please select at least one document to download.")
            else:
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                processed = 0
                for pdf_name, selected_urls in st.session_state["temp_selected_links"].items():
                    for url in selected_urls:
                        status_text.text(f"Downloading and analyzing {url.split('/')[-1]}...")
                        
                        fetched = fetch_external_document(url)
                        if fetched:
                            doc_name, doc_bytes, content_type = fetched
                            
                            # Store the downloaded file for later download
                            file_key = f"{pdf_name}||{url}"
                            st.session_state["downloaded_files"][file_key] = {
                                "name": doc_name,
                                "bytes": doc_bytes,
                                "mime": content_type or mimetypes.guess_type(doc_name)[0]
                            }
                            
                            # Analyze the document
                            result = analyze_document(
                                doc_name,
                                doc_bytes,
                                prompt_template,
                                active_api_key,
                                source_label=f"Downloaded from {pdf_name}",
                            )
                            
                            if result:
                                st.session_state["supporting_analysis_results"].append(result)
                        
                        processed += 1
                        progress_bar.progress(processed / total_selected)
                
                status_text.text("Analysis complete!")
                st.success(f"Analyzed {processed} supporting documents. Download buttons are now available.")
                st.rerun()
    
    # Display Step 2 Results
    if st.session_state.get("supporting_analysis_results"):
        st.write("---")
        st.subheader("Supporting Documents Analysis Results")
        
        supporting_df = pd.DataFrame(st.session_state["supporting_analysis_results"]).fillna("")
        supporting_df = order_dataframe_columns(supporting_df)
        
        st.dataframe(supporting_df, use_container_width=True)
        
        # Download combined results
        st.write("---")
        st.subheader("Download Complete Analysis")
        
        combined_excel = create_excel_output(
            st.session_state["primary_analysis_results"],
            st.session_state["supporting_analysis_results"]
        )
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.download_button(
                label="Download Combined Analysis (Excel)",
                data=combined_excel,
                file_name=f"complete_rfp_analysis_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        
        with col2:
            # Supporting docs only
            supporting_excel = create_excel_output([], st.session_state["supporting_analysis_results"])
            st.download_button(
                label="Download Supporting Docs Only (Excel)",
                data=supporting_excel,
                file_name=f"supporting_docs_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        
        with col3:
            # Combined CSV
            all_results = st.session_state["primary_analysis_results"] + st.session_state["supporting_analysis_results"]
            combined_df = pd.DataFrame(all_results).fillna("")
            combined_df = order_dataframe_columns(combined_df)
            csv_data = combined_df.to_csv(index=False)
            st.download_button(
                label="Download Combined Analysis (CSV)",
                data=csv_data,
                file_name=f"complete_rfp_analysis_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
            )

# Reset button
if st.session_state["primary_analysis_results"] or st.session_state.get("supporting_analysis_results"):
    st.write("---")
    if st.button("Start New Analysis"):
        st.session_state["primary_analysis_results"] = []
        st.session_state["supporting_analysis_results"] = []
        st.session_state["extracted_links"] = {}
        st.session_state["attachment_cache"] = {}
        st.session_state["temp_selected_links"] = {}
        st.session_state["current_step"] = 1
        st.rerun()

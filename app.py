import streamlit as st
import pandas as pd
import fitz  # PyMuPDF
from io import BytesIO
from google import genai
import json

# --- Page Configuration ---
st.set_page_config(
    page_title="RFP Analysis Tool",
    page_icon="📄",
    layout="wide"
)

st.title("📄 RFP & SOW Analysis Automation Tool")
st.write("Upload government tender PDFs (RFP or SOW documents) to automatically extract key information.")

# --- API Key Configuration ---
with st.sidebar:
    st.header("⚙️ Configuration")
    api_key = st.text_input("Enter your Google Gemini API Key", type="password")
    st.caption("Don't have a key? [Get one from Google AI Studio](https://aistudio.google.com/)")
    
    st.divider()
    
    st.subheader("Model Selection")
    model_choice = st.radio(
        "Choose AI Model:",
        ["gemini-2.5-flash (Faster)", "gemini-2.5-pro (More Accurate)"],
        help="Flash is faster but Pro gives better extraction quality"
    )
    selected_model = "gemini-2.5-flash" if "flash" in model_choice else "gemini-2.5-pro"
    
    st.divider()
    
    st.subheader("📋 Extracted Fields")
    with st.expander("RFP Fields"):
        st.markdown("""
        - Bid Number & Dates
        - Ministry/Department/Office
        - Estimated Value & Contract Period
        - Turnover & Experience Requirements
        - MSE/Startup Exemptions
        - Scope of Work Summary
        - Go/No-Go Eligibility Summary
        - Linked Document URLs
        """)
    
    with st.expander("SOW Fields"):
        st.markdown("""
        - Technical Requirements
        - Deliverables & Timeline
        - Performance Metrics
        - Quality Standards
        - Implementation Methodology
        - Support & Maintenance
        """)

# --- Core Functions ---

def extract_text_and_urls_from_pdf(pdf_file):
    """Extract text content and hyperlinks from PDF"""
    try:
        pdf_bytes = pdf_file.read()
        pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")
        
        # Extract text
        text = ""
        for page in pdf_document:
            text += page.get_text()
        
        # Extract URLs
        urls = []
        for page_num in range(len(pdf_document)):
            page = pdf_document[page_num]
            links = page.get_links()
            for link in links:
                if 'uri' in link:
                    uri = link['uri']
                    if uri and (uri.startswith('http://') or uri.startswith('https://')):
                        if uri not in urls:
                            urls.append(uri)
        
        pdf_document.close()
        return text, urls
    except Exception as e:
        st.error(f"Error reading PDF {pdf_file.name}: {e}")
        return None, []

def extract_rfp_data(text_content, filename, urls):
    """Extract RFP data using AI"""
    urls_string = "; ".join(urls) if urls else "No linked documents found"
    
    prompt = f"""
You are an expert Government Tender Analyst. Extract data from this GeM tender document into a JSON object.

Tender Document: {filename}
---
{text_content[:50000]}
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
    
    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model=selected_model,
            contents=prompt
        )
        
        # Parse JSON response
        json_string = response.text.strip().replace('```json', '').replace('```', '').strip()
        data = json.loads(json_string)
        return data
    except json.JSONDecodeError as e:
        st.error(f"Failed to parse AI response as JSON: {e}")
        return None
    except Exception as e:
        st.error(f"AI Error: {e}")
        return None

def extract_sow_data(text_content, filename):
    """Extract SOW data using AI"""
    prompt = f"""
You are an expert Government Tender Analyst. Extract SOW data from this document into a JSON object.

Document: {filename}
---
{text_content[:50000]}
---

**Instructions:**
1. Extract all technical and operational details
2. Use "Not Found" if information is missing
3. Provide comprehensive summaries for complex fields

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
    
    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model=selected_model,
            contents=prompt
        )
        
        json_string = response.text.strip().replace('```json', '').replace('```', '').strip()
        data = json.loads(json_string)
        data["Document_Type"] = "SOW"
        return data
    except json.JSONDecodeError as e:
        st.error(f"Failed to parse AI response as JSON: {e}")
        return None
    except Exception as e:
        st.error(f"AI Error: {e}")
        return None

# --- Main UI ---

uploaded_files = st.file_uploader(
    "📁 Choose PDF files", 
    type="pdf", 
    accept_multiple_files=True,
    help="Upload one or more RFP or SOW documents"
)

if uploaded_files:
    st.subheader("Document Type Selection")
    st.caption("Select the document type for each uploaded file")
    
    doc_types = {}
    cols = st.columns(min(len(uploaded_files), 3))
    
    for idx, file in enumerate(uploaded_files):
        with cols[idx % 3]:
            doc_types[file.name] = st.radio(
                file.name,
                options=["RFP", "SOW"],
                key=f"type_{idx}",
                horizontal=True
            )

if st.button("🚀 Analyze Documents", type="primary", disabled=not uploaded_files):
    if not api_key:
        st.warning("⚠️ Please enter your Gemini API Key in the sidebar to proceed.")
    elif not uploaded_files:
        st.warning("⚠️ Please upload at least one PDF file.")
    else:
        rfp_results = []
        sow_results = []
        
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, pdf_file in enumerate(uploaded_files):
            status_text.info(f"📄 Processing: **{pdf_file.name}** ({idx+1}/{len(uploaded_files)})")
            
            text, urls = extract_text_and_urls_from_pdf(pdf_file)
            
            if text:
                doc_type = doc_types[pdf_file.name]
                
                if doc_type == "RFP":
                    data = extract_rfp_data(text, pdf_file.name, urls)
                    if data:
                        rfp_results.append(data)
                        st.success(f"✅ Successfully analyzed {pdf_file.name} (RFP)")
                else:  # SOW
                    data = extract_sow_data(text, pdf_file.name)
                    if data:
                        sow_results.append(data)
                        st.success(f"✅ Successfully analyzed {pdf_file.name} (SOW)")
            
            progress_bar.progress((idx + 1) / len(uploaded_files))
        
        status_text.success("✨ Analysis Complete!")
        
        # Display Results
        if rfp_results or sow_results:
            st.divider()
            st.header("📊 Analysis Results")
            
            # Create Excel file
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                if rfp_results:
                    df_rfp = pd.DataFrame(rfp_results)
                    df_rfp.to_excel(writer, index=False, sheet_name='RFP_Data')
                    
                    st.subheader("RFP Documents")
                    st.dataframe(df_rfp, use_container_width=True)
                
                if sow_results:
                    df_sow = pd.DataFrame(sow_results)
                    df_sow.to_excel(writer, index=False, sheet_name='SOW_Data')
                    
                    st.subheader("SOW Documents")
                    st.dataframe(df_sow, use_container_width=True)
            
            excel_data = output.getvalue()
            
            st.download_button(
                label="📥 Download Complete Analysis as Excel",
                data=excel_data,
                file_name="rfp_sow_analysis_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        else:
            st.error("❌ No data was successfully extracted. Please check your documents and try again.")

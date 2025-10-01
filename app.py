# app.py - RFP Analyzer using google.genai Client

import streamlit as st
import pandas as pd
import fitz # PyMuPDF
from io import BytesIO
from google import genai

# --- Page Configuration ---
st.set_page_config(
    page_title="RFP Analysis Tool",
    page_icon="📄",
    layout="wide"
)

st.title("📄 RFP Analysis Automation Tool")
st.write("Upload one or more government tender PDFs to automatically extract key information.")

# --- API Key Configuration ---
with st.sidebar:
    st.header("Configuration")
    api_key = st.text_input("Enter your Google Gemini API Key", type="password")
    st.caption("Don't have a key? [Get one from Google AI Studio](https://aistudio.google.com/)")

# --- Core Functions ---

def extract_text_from_pdf(pdf_file):
    try:
        pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
        text = "".join(page.get_text() for page in pdf_document)
        pdf_document.close()
        return text
    except Exception as e:
        st.error(f"Error reading PDF {pdf_file.name}: {e}")
        return None

def get_data_from_ai(text_content):
    prompt = f"""
Based on the text from the RFP document provided below, extract the following information.
If a specific piece of information is not found, please explicitly state 'Not Found'.

**Bid Details:**
- Bid Number:
- Bid End Date/Time:
- Bid Opening Date/Time:
- Estimated Bid Value:
- Bid Offer Validity (From End Date):

**Department and Location:**
- Ministry/State Name:
- Department Name:
- Organisation Name:
- Office Name:

**Contract Details:**
- Contract Period:
- Item Category/Type:

**Scope:**
- Scope of work: (Provide a brief summary)

**Eligibility Criteria:**
- Minimum Average Annual Turnover of the bidder:
- Years of Past Experience Required for same/similar service:
- Past Experience of Similar Services required: (Summarize the requirement)
- MSE Exemption for Years Of Experience and Turnover:
- Startup Exemption for Years Of Experience and Turnover:

Here is the RFP text:
---
{text_content}
---
"""
    try:
        client = genai.Client(api_key=api_key)
        response = client.models.generate_content(
            model="models/gemini-2.5-pro",
            contents=prompt
        )
        return response.text
    except Exception as e:
        return f"Error communicating with AI: {e}"

def parse_ai_response(response_text):
    data_dict = {}
    lines = response_text.strip().split('\n')
    for line in lines:
        if ':' in line:
            parts = line.split(':', 1)
            key = parts[0].strip().replace("-", "").strip()
            value = parts[1].strip()
            data_dict[key] = value
    return data_dict

# --- Streamlit UI Elements ---
uploaded_files = st.file_uploader(
    "Choose RFP PDF files", 
    type="pdf", 
    accept_multiple_files=True
)

if st.button("🚀 Analyze RFPs"):
    if not api_key:
        st.warning("Please enter your Gemini API Key in the sidebar to proceed.")
    elif not uploaded_files:
        st.warning("Please upload at least one PDF file.")
    else:
        with st.spinner('Analyzing documents... Please wait.'):
            all_results = []

            for pdf_file in uploaded_files:
                st.write("---")
                st.info(f"Processing: **{pdf_file.name}**")

                text = extract_text_from_pdf(pdf_file)

                if text:
                    ai_response = get_data_from_ai(text)
                    if "Error" not in ai_response:
                        parsed_data = parse_ai_response(ai_response)
                        parsed_data['File Name'] = pdf_file.name
                        all_results.append(parsed_data)
                        st.success(f"Successfully analyzed {pdf_file.name}")
                    else:
                        st.error(f"Failed to analyze {pdf_file.name}: {ai_response}")

            if all_results:
                st.write("---")
                st.header("Analysis Results")

                df = pd.DataFrame(all_results)

                if 'File Name' in df.columns:
                    cols = ['File Name'] + [col for col in df.columns if col != 'File Name']
                    df = df[cols]

                st.dataframe(df)

                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='RFP_Data')

                excel_data = output.getvalue()

                st.download_button(
                    label="📥 Download Data as Excel",
                    data=excel_data,
                    file_name="rfp_analysis_results.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
import os
import io
import streamlit as st
import openai
from openai import OpenAI
import openpyxl
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
from dotenv import load_dotenv
from tempfile import NamedTemporaryFile

load_dotenv()

AZURE_DOCUMENT_ANALYZER_KEY = os.getenv("AZURE_DOCUMENT_ANALYZER_KEY")
AZURE_DOCUMENT_ANALYZER_ENDPOINT = os.getenv("AZURE_DOCUMENT_ANALYZER_ENDPOINT")

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

document_client = DocumentAnalysisClient(
    endpoint=AZURE_DOCUMENT_ANALYZER_ENDPOINT,
    credential=AzureKeyCredential(AZURE_DOCUMENT_ANALYZER_KEY)
)

openai_client = OpenAI(api_key=OPENAI_API_KEY)

st.set_page_config(
    page_title="Financial Statement Analyzer",
    page_icon="💼",
    layout="centered"
)

st.title("💼 Financial Statement Analyzer")
st.markdown("""
Upload a PDF file of a financial statement, and this demo app will extract key financial metrics for you.
""")

uploaded_file = st.file_uploader("Upload Financial Statement PDF", type=["pdf"])

if uploaded_file is not None:
    with st.spinner('Processing... This may take a few minutes.'):
        try:
            with NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf.write(uploaded_file.read())
                pdf_path = temp_pdf.name

            '''
            This function takes in a PDF file path, extracts tables from the PDF
            using Azure Document Analysis and returns the tables as a list of rows.
            '''
            def extract_tables_from_pdf(pdf_path):
                try:
                    with open(pdf_path, "rb") as pdf_file:
                        poller = document_client.begin_analyze_document("prebuilt-layout", document=pdf_file)
                        result = poller.result()

                    tables = []
                    for table in result.tables:
                        rows = []
                        for row_idx in range(table.row_count):
                            row = []
                            for cell in table.cells:
                                if cell.row_index == row_idx:
                                    row.append(cell.content)
                            rows.append(row)
                        tables.append(rows)
                    return tables
                except Exception as e:
                    print(f"Error extracting tables: {e}")
                    return []

            '''
            This function takes in the parsed tables and prompts OpenAI
            to extract the desired information and format it. 
            '''
            def parse_tables_with_openai(tables):
                with open('data/prompt.txt', 'r') as f:
                    prompt = f.read()
                
                for table in tables:
                    prompt += "\n".join(["\t".join(row) for row in table]) + "\n\n"

                try:
                    response = openai_client.chat.completions.create(
                        model="gpt-4-turbo", 
                        messages=[
                            {"role": "system", "content": "You are a helpful assistant for financial analysis."},
                            {"role": "user", "content": prompt}
                        ]
                    )
                    return response.choices[0].message.content
                except Exception as e:
                    print(f"Error parsing tables with OpenAI: {e}")
                    return ""

            '''
            This function takes the extracted information and 
            saves it to an excel file. 
            '''
            def save_to_excel(parsed_data):
                try:
                    workbook = openpyxl.Workbook()
                    sheet = workbook.active
                    sheet.title = "Financial Metrics"

                    lines = parsed_data.strip().split("\n")
                    for line in lines:
                        cells = [cell.strip() for cell in line.split("\t")]
                        sheet.append(cells)

                    # save to a BytesIO object
                    excel_file = io.BytesIO()
                    workbook.save(excel_file)
                    excel_file.seek(0)
                    return excel_file
                except Exception as e:
                    st.error(f"Error saving data to Excel: {e}")
                    return None

            # call helper functions to parse everything
            tables = extract_tables_from_pdf(pdf_path)
            if not tables:
                st.error("No tables found in the PDF.")
            else:
                parsed_data = parse_tables_with_openai(tables)
                if parsed_data:
                    excel_file = save_to_excel(parsed_data)
                    if excel_file:
                        st.success("Processing complete!")
                        st.markdown("### Download Extracted Metrics")
                        st.download_button(
                            label="Download Excel File",
                            data=excel_file,
                            file_name='financial_metrics.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                    else:
                        st.error("Failed to create Excel file.")
                else:
                    st.error("Failed to parse data with OpenAI.")
        except Exception as e:
            st.error(f"An error occurred: {e}")
else:
    st.info("Please upload a PDF file to get started.")

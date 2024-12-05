'''
I wrote a streamlit app to extract financial metrics from a PDF file using Docling document analysis and
OpenAI's GPT-4 model. Components are composable so the user can swap out various parts of the pipeline (i.e.
use Azure Document Intelligence instead of Docling, or use a different OpenAI model). 

This app has a user-friendly interface and is easy to deploy.
'''

import openai
from openai import OpenAI
import streamlit as st
import os
import openpyxl
from prompt import FinancialStatementExtract
import json
from docling.document_converter import DocumentConverter
from dotenv import load_dotenv
from tempfile import NamedTemporaryFile
from azure.core.credentials import AzureKeyCredential
import io

load_dotenv()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

openai_client = OpenAI(api_key=OPENAI_API_KEY)

st.set_page_config(
    page_title="Financial Statement Analyzer",
    page_icon="💼",
    layout="centered"
)

st.title("💼 Financial Statement Analyzer")
st.markdown("""
Upload a PDF file of a financial statement, and this demo will extract key financial metrics for you.
""")

uploaded_file = st.file_uploader("Upload Financial Statement PDF", type=["pdf"])

if uploaded_file is not None:
    with st.spinner('Processing...'):
        try:
            with NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf.write(uploaded_file.read())
                pdf_path = temp_pdf.name

            
            def extract_tables_from_pdf(pdf_path):
                try:
                    converter = DocumentConverter() # Uses Docling for now. 
                    result = converter.convert(pdf_path)
                    tables = result.document.export_to_markdown()
                    print(tables)
                    return tables
                except Exception as e:
                    print(f"Error extracting tables: {e}")
                    return []

            
            def parse_tables_with_openai(tables):
                with open('data/prompt2.txt', 'r') as f:
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

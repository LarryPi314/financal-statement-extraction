'''
This script processes a financial statement PDF file and extracts structured data using Azure Document Intelligence
and openAI's GPT-4 model.
'''

import openai
from openai import OpenAI
import os
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
import openpyxl
from prompt import FinancialStatementExtract
import json

from dotenv import load_dotenv

load_dotenv()

AZURE_DOCUMENT_ANALYZER_KEY = os.getenv("AZURE_DOCUMENT_ANALYZER_KEY")
AZURE_DOCUMENT_ANALYZER_ENDPOINT = os.getenv("AZURE_DOCUMENT_ANALYZER_ENDPOINT")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

document_client = DocumentAnalysisClient(
    endpoint=AZURE_DOCUMENT_ANALYZER_ENDPOINT,
    credential=AzureKeyCredential(AZURE_DOCUMENT_ANALYZER_KEY)
)
openai_client = OpenAI(api_key=OPENAI_API_KEY)

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
        response = openai_client.beta.chat.completions.parse(
            model="gpt-4o-2024-08-06", 
            messages=[
                {"role": "system", "content": "You are a helpful assistant for structured financial data extraction. You will be given unstructured text from a research paper and should convert it into the given structure."},
                {"role": "user", "content": prompt}
            ],
            response_format=FinancialStatementExtract
        )
        return response.choices[0].message.parsed
    except Exception as e:
        print(f"Error parsing tables with OpenAI: {e}")
        return ""

'''
This function takes the extracted information and 
saves it to an excel file. 
'''
def save_to_excel(parsed_data, output_path):
    try:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Financial Metrics"
        sheet.append(["Metric", "Value"])

        for key, value in parsed_data: # write parsed data to excel
            vals = value.split('|')
            vals.append('N/A')

            if value == 'N/A':
                sheet.append([str(key), 'N/A', 'N/A'])
            else:
                sheet.append([str(key), vals[0], vals[1]])

        workbook.save(output_path)
    except Exception as e:
        print(f"Error saving data to Excel: {e}")

'''
This function processes the financial statement PDF file
and invokes the above functionality. 
'''
def process_financial_statement(pdf_path, output_path):
    print("Extracting tables from PDF...")
    tables = extract_tables_from_pdf(pdf_path)

    print("Parsing tables with OpenAI...")
    parsed_data = parse_tables_with_openai(tables)

    print("Saving data to Excel...")
    save_to_excel(parsed_data, output_path)
    print(f"Processing complete. Data saved to {output_path}")

if __name__ == "__main__":
    pdf_path = "data/apple_financial_statement.pdf"  
    output_path = "data/apple_financial_metrics.xlsx"
    
    process_financial_statement(pdf_path, output_path)

'''
This Flask app hosts an endpoint that processes a financial statement PDF file and extracts structured data 
using Azure Document Intelligence and OpenAI's GPT-4 model. I wrote a HTTP endpoint to scale my app in production.
Components inside the pipeline can be swapped at will. 
'''

import openai
from openai import OpenAI
import os
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.core.credentials import AzureKeyCredential
import openpyxl
from dotenv import load_dotenv

from flask import Flask, request, jsonify, send_file
import tempfile

load_dotenv()

AZURE_DOCUMENT_ANALYZER_KEY = os.getenv("AZURE_DOCUMENT_ANALYZER_KEY")
AZURE_DOCUMENT_ANALYZER_ENDPOINT = os.getenv("AZURE_DOCUMENT_ANALYZER_ENDPOINT")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

document_client = DocumentAnalysisClient(
    endpoint=AZURE_DOCUMENT_ANALYZER_ENDPOINT,
    credential=AzureKeyCredential(AZURE_DOCUMENT_ANALYZER_KEY)
)

openai_client = OpenAI(api_key=OPENAI_API_KEY)

app = Flask(__name__)

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
This function takes in the extracted tables in markdown format
and prompts OpenAI to extract the desired information and format it.
'''
def parse_tables_with_openai(tables_markdown):
    with open('data/prompt.txt', 'r') as f:
        prompt = f.read()

    # Append the extracted tables to the prompt
    for table_markdown in tables_markdown:
        prompt += "\n" + table_markdown + "\n"

    try:
        response = openai_client.chat.completions.create(
            model="gpt-4",
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
saves it to an Excel file.
'''
def save_to_excel(parsed_data, output_path):
    try:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Financial Metrics"

        sheet.append(["Metric", "Value"])

        for line in parsed_data.split("\n"):
            vals = line.split("|")
            if len(vals) >= 2:
                sheet.append([x.strip() for x in vals[:2]])
        workbook.save(output_path)
    except Exception as e:
        print(f"Error saving data to Excel: {e}")

'''
This function processes the financial statement PDF file
and invokes the above functionality.
'''
def process_financial_statement(pdf_path, output_path):
    print("Extracting tables from PDF using OpenAI Vision...")
    tables_markdown = extract_tables_from_pdf(pdf_path)

    print("Parsing tables with OpenAI...")
    parsed_data = parse_tables_with_openai(tables_markdown)
    print(parsed_data)

    print("Saving data to Excel...")
    save_to_excel(parsed_data, output_path)
    print(f"Processing complete. Data saved to {output_path}")

@app.route('/process_financial_statement', methods=['POST'])
def process_financial_statement_endpoint():

    if 'file' not in request.files: # user needs to submit a file header
        return jsonify({'error': 'No file part in the request'}), 400

    file = request.files['file']

    # If the user does not select a file
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400

    if file:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            file.save(temp_pdf.name)
            pdf_path = temp_pdf.name

        # Create a temporary file for the output Excel file
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_xlsx:
            output_path = temp_xlsx.name

        # Call the process_financial_statement function
        process_financial_statement(pdf_path, output_path)

        # Send the output Excel file back to the user
        return send_file(output_path, as_attachment=True, attachment_filename='financial_metrics.xlsx')

    return jsonify({'error': 'File processing failed'}), 500

if __name__ == "__main__":
    app.run(debug=True)

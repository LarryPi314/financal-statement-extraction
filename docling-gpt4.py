'''
This script processes a financial statement PDF file and extracts structured data using Docling's document analysis
and openAI's GPT-4 model.
'''

import openai
from openai import OpenAI
import os
import openpyxl
from prompt import FinancialStatementExtract
import json
from docling.document_converter import DocumentConverter
from dotenv import load_dotenv

load_dotenv()

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

openai_client = OpenAI(api_key=OPENAI_API_KEY)

'''
This function takes in a PDF file path, extracts tables from the PDF
using Docling's document analysis and returns the tables as a list of rows.
'''
def extract_tables_from_pdf(pdf_path):
    try:
        converter = DocumentConverter()
        result = converter.convert(pdf_path)
        tables = result.document.export_to_markdown()
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
    
    prompt += f"\n\n{tables}"
    print(prompt)
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
        sheet.append(["Metric", "Year 2023", "Year 2024"])

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
    print(parsed_data)

    print("Saving data to Excel...")
    save_to_excel(parsed_data, output_path)
    print(f"Processing complete. Data saved to {output_path}")

if __name__ == "__main__":
    pdf_path = "data/apple_2.pdf"  
    output_path = "data/apple_financial_metrics_2.xlsx"
    
    process_financial_statement(pdf_path, output_path)

import json
import re

import fitz
from dotenv import load_dotenv
from openai import AzureOpenAI
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from docling.datamodel.base_models import InputFormat
from docling.document_converter import DocumentConverter, PdfFormatOption
from docling.datamodel.pipeline_options import PdfPipelineOptions, TableFormerMode
from docling.datamodel.settings import settings
import pandas as pd

load_dotenv()
client = AzureOpenAI(
    api_key="c3ec2f1",
    api_version="2025-01-01-preview",
    azure_endpoint="https://crisil-gen-ai-uat.openai.azure.com"
)

PDF_PATH = r"C:\Users\DeshmukhK\Downloads\12th Annual Report 2023-24.pdf"
OUTPUT_PDF_PATH = r"C:\Users\DeshmukhK\Downloads\12th Annual Report 2023-24_new.pdf"


def load_pdf_pages(pdf_path):
    """
    Load a PDF file and return its content as a list of strings, each representing a page.
    """
    pdf_document = fitz.open(pdf_path)
    pages = []
    for page in range(len(pdf_document)):
        text = pdf_document[page].get_text("text")
        pages.append({"page_num": page, "text": text})
    return pages, pdf_document


def keyword_prefilter(pages):
    """
    Filter pages based on the presence of specific keywords.
    """
    pattern = re.compile(r"\bcontingent\s+liabilit(y|ies)\b", re.IGNORECASE)
    return [p for p in pages if pattern.search(p['text'])]


def stage_1_classify(page_text):
    """
    Classify the page text using a pre-trained model.
    """
    prompt = f"""
    You are an expert Financial Analyst. You will be given the text content of a PDF page from an annual report.
    Determine if the page contains the "Contingent Liabilities" table.
    
    Rules:
    1. The page must contain an actual table, not just a reference.
    2. Ignore partial or incomplete mentions.
    
    Return JSON only:
    {{
      "relevance": "Relevant" or "Non Relevant"
      "confidence": float between 0 and 1
    }}
    Page Text:
    {page_text}
    """
    # resp = client.chat.completions.create(input=prompt, temperature=0)
    resp = client.chat.completions.create(model="gpt-4.1-mini", temperature=0, max_tokens=16000, seed=42,
                                          messages=[{"role": "user", "content": prompt}],
                                          response_format={"type": "json_object"})

    return resp.choices[0].message.content.strip()


def classifyTable(table_makrdown: str = ""):
    prompt = f"""
    You are a highly reliable AI assistant with deep expertise in financial reporting and analysis, especially in identifying tables related to contingent liabilities.

    **Task:**  
    Given the content of a financial table, determine if it represents a contingent liabilities table.
    
    **Instructions:**  
    - Carefully analyze the table content provided below.
    - Look for keywords or phrases commonly associated with contingent liabilities, such as:  
      - "contingent liability", "contingent liabilities", "guarantees", "claims", "litigation", "disputed", "not acknowledged as debt", "pending cases", "bonds", "letters of credit", "bank guarantees", "legal proceedings", "tax disputes", "unascertained liabilities", "claims against the company", "arbitration", "court cases", "surety", "indemnity", "commitments", "possible obligations"
    - Consider the context, column headers, and any notes or footnotes that indicate potential or possible obligations not yet recognized as actual liabilities.
    - Respond only with 'True' if the table is about contingent liabilities, or 'False' if it is not.
    - Do not provide explanations or any additional text.
    
    **Output:**  
    True or False
    
    
    Now generate your response for this table content: 
    **Table Content:**  
    {table_makrdown}
        """

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {"role": "system", "content": "Identify the if table is contigent liabilities or not"},
            {"role": "user", "content": prompt}
        ],
        temperature=0.1,
    )
    return response.choices[0].message.content


def stage_2_classify(page_text):
    """
    Classify the page text using a pre-trained model.
    """
    prompt = f"""
    You are verifying a page for accuracy.
    Confirm if the page contains the "COntingent Liabilities" table
    
    Requirements:
    - Must include the heading "Contingent Liabilities" oe clsoe variation.
    - Must have tabular format with multiple rows.
    - If any requirement is missing amrk as false.
    
    Return JSON only:
    {{
      "relevance": "Relevant" or "Non Relevant"
      "confidence": float between 0 and 1
    }}
    Page Text:
    {page_text}
    """
    # resp = client.responses.create(model="gpt-4.1-mini", message=prompt, temperature=0)
    resp = client.chat.completions.create(model="gpt-4.1-mini", temperature=0, max_tokens=16000, seed=42,
                                          messages=[{"role": "user", "content": prompt}],
                                          response_format={"type": "json_object"})
    return resp.choices[0].message.content.strip()


def parse_page_with_docling(pdf_path, page_num):
    """
    Parse a specific page of the PDF using Docling.
    """
    converter = DocumentConverter()
    result = converter.convert(pdf_path, page_range=(page_num, page_num + 1))
    return result.document.export_to_markdown()


def classifyTable(table_makrdown: str = ""):
    prompt = f"""
    You are a highly reliable AI assistant with deep expertise in financial reporting and analysis, especially in identifying tables related to contingent liabilities.

    **Task:**  
    Given the content of a financial table, determine if it represents a contingent liabilities table.
    
    **Instructions:**  
    - Carefully analyze the table content provided below.
    - Look for keywords or phrases commonly associated with contingent liabilities, such as:  
      - "contingent liability", "contingent liabilities", "guarantees", "claims", "litigation", "disputed", "not acknowledged as debt", "pending cases", "bonds", "letters of credit", "bank guarantees", "legal proceedings", "tax disputes", "unascertained liabilities", "claims against the company", "arbitration", "court cases", "surety", "indemnity", "commitments", "possible obligations"
    - Consider the context, column headers, and any notes or footnotes that indicate potential or possible obligations not yet recognized as actual liabilities.
    - Respond only with 'True' if the table is about contingent liabilities, or 'False' if it is not.
    - Do not provide explanations or any additional text.
    
    **Output:**  
    True or False
    
    
    Now generate your response for this table content: 
    **Table Content:**  
    {table_makrdown}
        """

    response = client.chat.completions.create(
        model="gpt-4.1-mini",
        messages=[
            {"role": "system", "content": "Identify the if table is contigent liabilities or not"},
            {"role": "user", "content": prompt}
        ],
        temperature=0.1,
    )

    return response.choices[0].message.content


def extract_table_from_docling_markdown(markdown_text):
    """
    Extract the Contingent Liabilities table from the Docling markdown text.
    """
    prompt = f"""
    You are a financial data extraction expert.
    You are given a markdown version of a PDF page with clearly formatted tables.
    
    Your task:
    - Identify the table titled "Contingent Liabilities" or close variations.
    - Extract ONLY that table into a structured JSON where each row is a dictionary.
    - Ignore all other tables.
    
    Return JSON only in the format:
    [
        {{ "Column1": "Value1", "Column2": "Value2", ... }},
        ...
    ]
    If no contingent Liabilities table is found, return [].
    
    Markdown:
    {markdown_text}
    """

    # resp = client.(model="gpt-4.1-mini", message=prompt, temperature=0)
    resp = client.chat.completions.create(model="gpt-4.1-mini", temperature=0, max_tokens=16000, seed=42,
                                          messages=[{"role": "user", "content": prompt}],
                                          response_format={"type": "json_object"})
    try:
        return json.loads(resp.choices[0].message.content)
    except json.JSONDecodeError:
        print(f"Error parsing JSON from response: {resp.output_text.strip()}")
        return []


def clean_illegal_chars(df):
    return df.applymap(
        lambda x: ILLEGAL_CHARACTERS_RE.sub("", str(x)) if isinstance(x, str) else x
    )

def get_docling_pipeline():
    """
    Get docling pipeline.

    """
    try:
        pipeline_options = PdfPipelineOptions(do_table_structure=True,
                                              table_structure_options=dict(do_cell_matching=True,mode=TableFormerMode.ACCURATE))

        # Initialize the DocumentConverter with the specified pipeline options
        doc_converter_global = DocumentConverter(format_options={InputFormat.PDF: PdfFormatOption(pipeline_options=pipeline_options)})

        settings.debug.profile_pipeline_timings = True
        return doc_converter_global

    except Exception as e:
        print(f"Exception occured while getting the docling pipeline :{e}")


doc_converter_global = get_docling_pipeline()

def get_docling_results(INPUT_PDF_SOURCE):
    result = doc_converter_global.convert(INPUT_PDF_SOURCE)
    return result
def extract_cg(pdf_path):
    """
    Main function to extract Contingent Liabilities from a PDF.
    """
    pages, pdf_document = load_pdf_pages(pdf_path)
    candidates = keyword_prefilter(pages)

    print(f"[INFO] Prefiltered {len(candidates)} pages containing 'Contingent Liabilities' keyword.")

    relevant_pages = []

    for p in candidates:
        stage1_result = stage_1_classify(p['text'])
        print(f"[DEBUG] Stage 1 - Page {p['page_num'] + 1}: {stage1_result}")

        if "Relevant" in stage1_result and "confidence" in stage1_result:
            print("Inside Stage 2")
            confidence_match = re.search(r'"confidence":\s*([0-9.]+)', stage1_result)
            confidence = float(confidence_match.group(1)) if confidence_match else 0
            if confidence >= 0.85:
                stage2_result = stage_2_classify(p['text'])
                print(f"[DEBUG] Stage 2 - Page {p['page_num'] + 1}: {stage2_result}")

                if '"relevance": "Relevant"' in stage2_result:
                    relevant_pages.append(p['page_num'])

    if relevant_pages:
        pdf = fitz.open(pdf_path)
        new_pdf = fitz.open()

        for page in relevant_pages:
            new_pdf.insert_pdf(pdf, from_page=page, to_page=page)
        new_pdf.save(OUTPUT_PDF_PATH)
        new_pdf.close()
        pdf.close()

        return relevant_pages
    else:
        return None
        print("No relevant pages found.")

def create_table(OUTPUT_PDF_PATH):




    # for pages in file['pages']:
    result = get_docling_results(INPUT_PDF_SOURCE=OUTPUT_PDF_PATH)
    # output_file_name = file.split("\\")[-1].replace(".pdf", "")
    previous_page_num = None
    current_page_table_count = 0
    for table_ix, table in enumerate(result.document.tables):
        current_page_num = table.dict()['prov'][0]['page_no']
        # print("current_page_num:", current_page_num)
        # print("previous_page_num :", previous_page_num)

        if previous_page_num is None:
            previous_page_num = current_page_num

        if previous_page_num == current_page_num:
            current_page_table_count += 1
        else:
            current_page_table_count = 1

        # print(current_page_table_count)

        sheet_name = f"Page_no_{current_page_num}_table_{current_page_table_count}"
        # print("sheet_name :", sheet_name)

        table_df: pd.DataFrame = table.export_to_dataframe()
        # Clean column names
        table_df.columns = [
            ILLEGAL_CHARACTERS_RE.sub("", str(col)) for col in table_df.columns
        ]
        table_df = clean_illegal_chars(table_df)  # Clean illegal characters
        # print(table_df)

        classification_result = classifyTable(table_df.to_markdown())

        print(classification_result)

        if "true" in classification_result.lower():
            # print(f"## Table {table_ix}")
            # print(table_df.to_markdown())

            # sheet_name = f"Table_{table_ix}"
            table_df.to_excel(sheet_name + ".xlsx", index=False)

        previous_page_num = current_page_num


extract_cg(PDF_PATH)
create_table(OUTPUT_PDF_PATH)








import json
import ast
from langchain.llms.base import LLM
from langchain.callbacks.manager import CallbackManagerForLLMRun
class HostedLLM(LLM):
    def __init__(self, endpoint:str, **kwargs):
        super().__init__(**kwargs)
    def _llm_type(self) -> str:
        return "Hosted LLM"
    def _call(self, prompt:str, stop=None, run_manager: CallbackManagerForLLMRun=None) -> str:
        try:
            prompt_template = f"""
<|begin_of_text|><|start_header_id|>user<|end_header_id|>
            {prompt}<|eot_id|><|start_header_id|>assistant<|end_header_id|>
            """
            payload = json.dumps({"provider": "tgi", "deployment": "Llama 3.3 v1", "spec_version": 1, "input_text": prompt_template, "params": { "temperature": 0.1}})
            headers = {'token': '0e57', 'Content-Type': 'application/json'}
            response = requests.request("POST", url="https://llmgateway.crisil.local/api/v1/llm", headers=headers, data=payload, verify=False)
            response_v = ast.literal_eval(response.text)
            resp_o = response_v['output']
            output = str(resp_o).replace(prompt_template, "")
            return output.strip()
        except Exception as e:
            return f"LLM Call Failed: {e}"

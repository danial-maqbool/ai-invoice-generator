Comprehensive Guide: Building an AI Invoice Generator with GitHub Models

This document outlines the step-by-step process for building an application that converts unstructured order text into a formatted invoice document.

Based on your requirements, this app will use GitHub Models to process the raw text and map it to a Word document template modeled after your reference file, invoice-INV-0070-Shaq-Jordan.pdf.

Phase 1: Architecture & Setup

We will use Python for the backend and AI processing, and Streamlit for a simple, interactive user interface.

Prerequisites

Python 3.8+ installed on your machine.

A GitHub Personal Access Token (PAT). You will need this to authenticate with GitHub Models.

Project Setup

Open your terminal and run the following commands to create your project folder and install the necessary libraries:

mkdir ai-invoice-generator
cd ai-invoice-generator
python -m venv venv

# On Mac/Linux:
source venv/bin/activate
# On Windows:
venv\Scripts\activate

# Install dependencies
pip install streamlit openai docxtpl python-dotenv


Note: We use the openai library because GitHub Models provides an OpenAI-compatible API endpoint.

Create an .env file in the root directory to store your GitHub token securely:

GITHUB_TOKEN=your_github_personal_access_token_here


Phase 2: Creating the Document Template

Your app cannot directly edit the invoice-INV-0070-Shaq-Jordan.pdf file easily. Instead, you need to recreate the visual layout of that PDF as a Microsoft Word (.docx) file.

Open Microsoft Word and design the invoice to look like invoice-INV-0070-Shaq-Jordan.pdf (with the "ML Trading International" header, tables, etc.).

Replace the specific text (like "Shaq Jordan" or "INV-0070") with Jinja2 placeholder tags.

Your Word document should look something like this in the text areas:

Header/Info Section:

Invoice Number: {{ invoice_number }}

Date: {{ invoice_date }}

Customer/Receiver: {{ customer_name }}

Address: {{ customer_address }}

Reseller: {{ reseller_info }}

Payment Status: {{ payment_status }}

Shipping Status: {{ shipping_status }}

Products Table:
Create a standard Word table with 5 columns (CODE, PRODUCT, QTY, UNIT PRICE, LINE TOTAL). The second row should contain this exact text to create a loop:

{% tr for item in items %}
{{ item.code }} | {{ item.product }} | {{ item.qty }} | {{ item.unit_price }} | {{ item.line_total }}
{% tr endfor %}


Footer:

Order Total: {{ order_total }}

Save this file as template.docx in your project folder.

Phase 3: Writing the Application Code

Create a file named app.py in your project folder. We will build this in three parts: AI Extraction, Document Generation, and the User Interface.

Part 1: AI Extraction using GitHub Models

Copy and paste this into app.py. This code configures the OpenAI SDK to route through GitHub Models and provides a highly specific prompt based on your sample data.

import os
import json
from datetime import datetime
import streamlit as st
from openai import OpenAI
from docxtpl import DocxTemplate
from dotenv import load_dotenv

load_dotenv()

# Configure client for GitHub Models
client = OpenAI(
    base_url="https://models.inference.ai.azure.com",
    api_key=os.environ["GITHUB_TOKEN"],
)

def extract_invoice_data(raw_text):
    prompt = f"""
    You are an AI assistant that extracts invoice data from raw text and formats it as a JSON object.
    
    Raw Text:
    "{raw_text}"
    
    Instructions:
    1. Extract customer name and address.
    2. Identify Reseller info (e.g., "Reseller J/L").
    3. Determine Payment Status (e.g., "Payment awaiting" -> "Awaiting Payment").
    4. Determine Shipping Status (e.g., "Ready for delivery" -> "Ready for Dispatch").
    5. Extract all line items. For each item:
       - Generate a short, logical product code (e.g., "Retatrutide" -> "RETA40", "GLOW Nasal" -> "GLOWN", "NAD Nasal" -> "NADN").
       - Format the product name clearly.
       - Extract quantity.
       - Extract unit price (format as string with currency symbol, e.g., "£30.00").
       - Calculate line total (format as string with currency symbol, e.g., "£2,100.00").
    6. Calculate the actual accurate Order Total based on the line items, ignoring mathematical errors in the raw text. Format as "£X,XXX.XX".
    7. Generate a random Invoice Number starting with INV- (e.g., INV-0071).
    8. Use today's date for invoice_date (DD/MM/YYYY).

    Return ONLY a valid JSON object with these exact keys:
    {{
        "invoice_number": "string",
        "invoice_date": "string",
        "customer_name": "string",
        "customer_address": "string",
        "reseller_info": "string",
        "payment_status": "string",
        "shipping_status": "string",
        "items": [
            {{
                "code": "string",
                "product": "string",
                "qty": "number",
                "unit_price": "string",
                "line_total": "string"
            }}
        ],
        "order_total": "string"
    }}
    """

    # Using GPT-4o via GitHub Models (you can also use Meta-Llama-3-70B-Instruct, etc.)
    response = client.chat.completions.create(
        model="gpt-4o", 
        messages=[{"role": "system", "content": prompt}],
        response_format={ "type": "json_object" },
        temperature=0.1
    )
    
    return json.loads(response.choices[0].message.content)


Part 2: Document Generation Logic

Add this function below the extraction code in app.py. It takes the JSON from GitHub Models and pushes it into your Word template.

def generate_document(data_dict, template_path, output_path):
    doc = DocxTemplate(template_path)
    
    # Handle multi-line addresses correctly in Word
    # Replace standard newlines with Word-friendly line breaks
    if "customer_address" in data_dict:
        data_dict["customer_address"] = data_dict["customer_address"].replace('\n', '\a')

    doc.render(data_dict)
    doc.save(output_path)
    return output_path


Part 3: Streamlit Web Interface

Add this to the bottom of app.py to create the user interface.

st.title("📄 AI Invoice Generator (GitHub Models)")

st.markdown(f"Upload your `template.docx` (based on **invoice-INV-0070-Shaq-Jordan.pdf**) and paste your raw order text.")

uploaded_template = st.file_uploader("Upload Word Template (.docx)", type="docx")

# Default text set to your sample input for easy testing
sample_input = """Reseller J/L
Shaq Jordan
Flat 21, 
46 falcon road 
SW11 2lr

70 × Retatrutide – £2,100 (£30 each)
10 × GLOW – £250 (£25 each)
10 × NAD Nasal – £250 (£25 each)
5 × GLOW Nasal – £100 (£20 each)
5 × Selank Nasal – £100 (£20 each)
5 × BPC/TB-500 Nasal – £100 (£20 each)
5 × BPC Nasal – £100 (£20 each)
5 × PT-141 Nasal – £100 (£20 each)
5 × Kisspeptin Nasal – £100 (£20 each)

Total: £3,300
(Payment awaiting) (Ready for delivery)"""

raw_text = st.text_area("Paste Raw Order Text", value=sample_input, height=300)

if st.button("Generate Invoice"):
    if not uploaded_template:
        st.error("Please upload your .docx template first.")
    else:
        with st.spinner("Processing with GitHub Models..."):
            # 1. Save uploaded template
            with open("temp_template.docx", "wb") as f:
                f.write(uploaded_template.getbuffer())
            
            try:
                # 2. Extract Data
                extracted_data = extract_invoice_data(raw_text)
                
                # Show the JSON on screen so you can verify the AI's work
                with st.expander("View Extracted Data"):
                    st.json(extracted_data)
                
                # 3. Generate Word Document
                output_file = f"Invoice_{extracted_data.get('customer_name', 'Generated').replace(' ', '_')}.docx"
                generate_document(extracted_data, "temp_template.docx", output_file)
                
                st.success("Invoice generated successfully!")
                
                # 4. Download Button
                with open(output_file, "rb") as file:
                    st.download_button(
                        label="Download Invoice (.docx)",
                        data=file,
                        file_name=output_file,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"An error occurred: {e}")


Phase 4: Running the Application

Make sure your .env file is saved with your GITHUB_TOKEN.

Make sure your template.docx is ready.

In your terminal, run:

streamlit run app.py


A web browser will open.

Upload your template.docx.

Click Generate Invoice.

The app will use GitHub Models to parse your exact input (correcting the mathematical error of £3,300 to £3,200 as seen in your reference PDF), generate the custom product codes (like RETA40), and map everything flawlessly into the downloaded .docx invoice.
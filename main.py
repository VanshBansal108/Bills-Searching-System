import os
import io
import openpyxl
from google.cloud import vision
from google.oauth2 import service_account
from datetime import datetime
from tqdm import tqdm  # progress bar
from google.cloud.vision_v1 import types

# Auth
credentials = service_account.Credentials.from_service_account_file("key.json")
client = vision.ImageAnnotatorClient(credentials=credentials)

# Excel setup
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Bill Data"
ws.append(["Year", "Seller", "Filename", "Invoice No", "Date", "Extracted Text"])

# Helper to extract invoice number & date from filename
def extract_info_from_filename(filename):
    try:
        base = os.path.splitext(filename)[0]
        parts = base.split("_")
        invoice = parts[0]
        date = datetime.strptime(parts[1], "%d-%b-%Y").date()
        return invoice, date.strftime("%Y-%m-%d")
    except:
        return "", ""

# Gather all files
file_list = []
for root, dirs, files in os.walk("bills"):
    for file in files:
        if file.lower().endswith((".jpg", ".jpeg", ".png", ".pdf")):
            file_list.append(os.path.join(root, file))

# Process files with progress bar
for filepath in tqdm(file_list, desc="Processing bills"):
    parts = filepath.split(os.sep)
    if len(parts) < 3:
        continue
    year = parts[-3]
    seller = parts[-2]
    file = os.path.basename(filepath)

    invoice_no, date_str = extract_info_from_filename(file)
    full_text = ""

    if file.lower().endswith(".pdf"):
        with io.open(filepath, 'rb') as pdf_file:
            content = pdf_file.read()
        mime_type = "application/pdf"
        input_config = types.InputConfig(content=content, mime_type=mime_type)

        request = types.AnnotateFileRequest(
            input_config=input_config,
            features=[types.Feature(type_=vision.Feature.Type.DOCUMENT_TEXT_DETECTION)]
        )

        responses = client.batch_annotate_files(requests=[request])
        for r in responses.responses:
            for resp in r.responses:
                if resp.full_text_annotation.text:
                    full_text += resp.full_text_annotation.text.replace("\n", " ") + " "
    else:
        with io.open(filepath, 'rb') as image_file:
            content = image_file.read()
        image = vision.Image(content=content)
        response = client.document_text_detection(image=image)
        full_text = response.full_text_annotation.text.strip().replace("\n", " ")

    ws.append([year, seller, file, invoice_no, date_str, full_text.strip()])

# Save Excel
wb.save("extracted_bills.xlsx")
print("✅ All done! Excel saved as extracted_bills.xlsx")

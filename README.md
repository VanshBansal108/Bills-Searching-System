Bills Searching System

This project automatically reads text from hundreds of bill PDFs or images and saves the results in an Excel file. 
It helps quickly find and organize bills using only simple information like company name, invoice number, amount paid, serial number of product, etc.

What It Does:
- Reads all the files inside the **bills/** folder (PDF, JPG, PNG).
- Uses **Google Vision API** to extract text from each bill.
- Collects the data (filename, date, invoice number, and full text).
- Saves everything neatly into **extracted_bills.xlsx**.

- Once the excel file is ready, the matter is copied onto a google sheet - which is linked with google drive
- This allows the user to find the relevant bill and open the pdf at anytime easily.


Tools Used
- Python  
- Google Cloud Vision API  
- openpyxl (for Excel)  
- ChatGPT (for guidance and debugging)
- Google sheets
- Google drive

About:

Built step-by-step with **ChatGPT**  
Developed at **Godlywood Studios, Mount Abu (2025)**

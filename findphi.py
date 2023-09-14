import os
import re
import docx
import openpyxl
from PyPDF2 import PdfReader
import warnings
import threading
import time

warnings.filterwarnings("ignore")

# Prompt the user for the path to search
network_share_path = input("Enter the path to search: ")

# Prompt the user for the path to save the report
output_file_path = input("Enter the path to save the report (e.g., C:\\BGInfo\\CPRA.txt): ")

# Define regular expressions for expected SSN and DOB formats
ssn_pattern = r"\b\d{3}-\d{2}-\d{4}\b"  # Matches XXX-XX-XXXX
dob_pattern = r"\b\d{2}/\d{2}/\d{4}\b"  # Matches XX/XX/XXXX

# Define a simple pattern to search for the word "password"
password_pattern = r"\bpassword\b"  # Matches the word "password"

search_phrases = [ssn_pattern, dob_pattern, password_pattern, "birthdate"]
file_extensions = ["docx", "xlsx", "pdf"]

# Function to get password with timeout
def get_password(timeout):
    password = None
    def input_thread():
        nonlocal password
        password = input(f"Enter the decryption password within {timeout} seconds (or press Enter to skip): ")

    input_thread = threading.Thread(target=input_thread)
    input_thread.daemon = True
    input_thread.start()
    input_thread.join(timeout)
    return password

# Use a context manager to open the output file for writing
with open(output_file_path, 'a') as f:
    for root, dirs, files in os.walk(network_share_path):
        for file in files:
            file_path = os.path.join(root, file)
            file_ext = file.split(".")[-1].lower()
            if file_ext in file_extensions:
                try:
                    # Skip temporary Word files
                    if file.startswith("~$"):
                        continue

                    if file_ext == "docx":
                        doc = docx.Document(file_path)
                        file_contents = "\n".join([para.text for para in doc.paragraphs])
                    elif file_ext == "xlsx":
                        wb = openpyxl.load_workbook(file_path)
                        sheet_names = wb.sheetnames
                        sheet_contents = []
                        for sheet_name in sheet_names:
                            sheet = wb[sheet_name]
                            for row in sheet.iter_rows():
                                for cell in row:
                                    if cell.value is not None:
                                        sheet_contents.append(str(cell.value))
                        file_contents = "\n".join(sheet_contents)
                    elif file_ext == "pdf":
                        pdf_file = open(file_path, 'rb')
                        reader = PdfReader(pdf_file)

                        # Check if the PDF is encrypted
                        if reader.is_encrypted:
                            # Get the password with a timeout
                            password = get_password(timeout=10)

                            if password:
                                if reader.decrypt(password):
                                    num_pages = len(reader.pages)
                                    page_contents = []
                                    for page_num in range(num_pages):
                                        page = reader.pages[page_num]
                                        page_text = page.extract_text()
                                        page_contents.append(page_text)
                                    file_contents = "\n".join(page_contents)
                                else:
                                    print(f"Failed to decrypt '{file_path}' with the provided password.")
                                    continue  # Skip the encrypted PDF file
                            else:
                                print(f"Skipped '{file_path}' due to timeout.")

                        else:
                            num_pages = len(reader.pages)
                            page_contents = []
                            for page_num in range(num_pages):
                                page = reader.pages[page_num]
                                page_text = page.extract_text()
                                page_contents.append(page_text)
                            file_contents = "\n".join(page_contents)

                    for search_phrase in search_phrases:
                        if isinstance(search_phrase, str):
                            # For non-regex search phrases, escape special characters
                            pattern = re.escape(search_phrase)
                        else:
                            pattern = search_phrase

                        matches = re.findall(pattern, file_contents, re.IGNORECASE)
                        if matches:
                            # Write the findings to the output file
                            f.write(f"Found '{search_phrase}' in file '{file_path}': {matches}\n")

                except Exception as e:
                    # Print detailed error message
                    print(f"Error reading '{file_path}': {e}")

print("Search and report generation completed.")

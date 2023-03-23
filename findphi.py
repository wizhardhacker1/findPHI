import os
import re
import docx
import openpyxl
import PyPDF2
import warnings

warnings.filterwarnings("ignore")

network_share_path = r"\\network\share\path"

search_phrases = ["ssn", "mrn", "dob", "password", "birthdate"]
ssn_pattern = r"\d{3}-\d{2}-\d{4}"
dob_pattern = r"\d{2}/\d{2}/\d{4}"

file_extensions = ["docx", "xlsx", "pdf"]

with open('c:\output.txt', 'a') as f:
    for root, dirs, files in os.walk(network_share_path):
        for file in files:
            file_path = os.path.join(root, file)
            file_ext = file.split(".")[-1].lower()
            if file_ext in file_extensions:
                try:
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
                        reader = PyPDF2.PdfFileReader(pdf_file)
                        num_pages = reader.getNumPages()
                        page_contents = []
                        for page_num in range(num_pages):
                            page = reader.getPage(page_num)
                            page_text = page.extractText()
                            page_contents.append(page_text)
                        file_contents = "\n".join(page_contents)

                    for search_phrase in search_phrases:
                        pattern = None
                        if search_phrase == "ssn":
                            pattern = ssn_pattern
                        elif search_phrase == "dob":
                            pattern = dob_pattern
                        else:
                            pattern = re.escape(search_phrase)

                        matches = re.findall(pattern, file_contents, re.IGNORECASE)
                        if matches:
                            warnings.simplefilter("ignore")
                            print(f"Found '{search_phrase}' in file '{file_path}': {matches}", file=f)

                except Exception as e:
                    print(f"Error reading '{file_path}': {e}")

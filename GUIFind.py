import os
import docx
import openpyxl
import re
from PyPDF2 import PdfReader
import warnings
import threading
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar
import html

warnings.filterwarnings("ignore")

# Function to get password with timeout
def get_password(timeout):
    password = None
    def input_thread():
        nonlocal password
        password = entry_password.get()

    input_thread = threading.Thread(target=input_thread)
    input_thread.daemon = True
    input_thread.start()
    input_thread.join(timeout)
    return password

# Function to search for patterns within words using regular expressions
def search_within_words(text, patterns):
    found = []
    for pattern, label in patterns:
        for match in re.finditer(pattern, text, re.IGNORECASE):
            found.append((label, match.group(0)))
    return found

# Function to search and report
def search_and_report():
    network_share_path = entry_search_path.get()
    output_file_path = entry_report_path.get()

    # Disable the Search button and set the message
    button_search.config(state=tk.DISABLED)
    label_status.config(text="Searching, please wait...")

    search_patterns = [
        (r'\d{3}-\d{2}-\d{4}', "Possible SSN"),      # SSN-like pattern
        (r'\d{2}/\d{2}/\d{4}', "Possible DOB"),     # DOB-like pattern (XX/XX/XXXX)
        (r'\d{4}-\d{2}-\d{2}', "Possible DOB"),     # DOB-like pattern (XX-XX-XXXX)
        (r'[A-Za-z\d@$!%*?&]{8,}', "Possible Password")  # Password-like pattern
    ]

    found_results = {
        "Possible SSN": [],
        "Possible DOB": [],
        "Possible Password": []
    }

    for root, dirs, files in os.walk(network_share_path):
        for file in files:
            file_path = os.path.join(root, file)
            file_ext = file.split(".")[-1].lower()
            if file_ext in ["docx", "xlsx", "pdf"]:
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
                            # Skip encrypted PDFs automatically
                            continue

                        num_pages = len(reader.pages)
                        page_contents = []
                        for page_num in range(num_pages):
                            page = reader.pages[page_num]
                            page_text = page.extract_text()
                            page_contents.append(page_text)
                        file_contents = "\n".join(page_contents)

                    # Search for patterns within words
                    found_patterns = search_within_words(file_contents, search_patterns)

                    # Append the findings to the appropriate column
                    for label, matches in found_patterns:
                        if matches:
                            found_results[label].append((html.escape(file_path), html.escape(matches)))

                except Exception as e:
                    # Print detailed error message
                    print(f"Error reading '{file_path}': {e}")

            # Update the progress bar
            progress_value.set((files.index(file) + 1) / len(files) * 100)
            app.update_idletasks()

    # Enable the Search button and set the completion message
    button_search.config(state=tk.NORMAL)
    label_status.config(text="Search completed.")
    progress_value.set(0)  # Clear the progress bar

    # Generate the HTML report
    generate_html_report(output_file_path, found_results)

    # Display a message box indicating the search is complete
    messagebox.showinfo("Search Completed", "Search and report generation completed. Results saved to the output file.")

def generate_html_report(output_file_path, found_results):
    with open(output_file_path, 'w') as f:
        f.write("<html><head><title>PHI Search Results</title></head><body>")
        f.write("<h1>PHI Search Results</h1>")

        # Create separate tables for each column
        for label, results in found_results.items():
            f.write(f"<h2>{label}</h2>")
            f.write("<table border='1' cellspacing='0' cellpadding='5'><tr><th>File Path</th><th>Matches</th></tr>")
            for file_path, matches in results:
                f.write(f"<tr><td>{file_path}</td><td>{matches}</td></tr>")
            f.write("</table>")

        f.write("</body></html>")

# Create the main application window
app = tk.Tk()
app.title("File Search and Reporting")

# Create and configure labels and entry fields
label_search_path = tk.Label(app, text="Enter the path to search:")
entry_search_path = tk.Entry(app)
label_report_path = tk.Label(app, text="Enter the path/name to save the HTML report: ex report.html")
entry_report_path = tk.Entry(app)
label_password = tk.Label(app, text="Enter the decryption password (optional):")
entry_password = tk.Entry(app, show="*")  # Password entry field

# Create and configure buttons
button_search = tk.Button(app, text="Search and Report", command=lambda: threading.Thread(target=search_and_report).start())
button_browse_search = tk.Button(app, text="Browse", command=lambda: entry_search_path.insert(0, filedialog.askdirectory()))
button_browse_report = tk.Button(app, text="Browse", command=lambda: entry_report_path.insert(0, filedialog.asksaveasfilename(defaultextension=".html")))

# Create and configure a label for status messages
label_status = tk.Label(app, text="")

# Create a progress bar
progress_value = tk.DoubleVar()
progress_bar = Progressbar(app, mode="determinate", variable=progress_value)

# Organize widgets using the grid layout
label_search_path.grid(row=0, column=0, sticky="e")
entry_search_path.grid(row=0, column=1, columnspan=2, sticky="ew")
button_browse_search.grid(row=0, column=3)
label_report_path.grid(row=1, column=0, sticky="e")
entry_report_path.grid(row=1, column=1, columnspan=2, sticky="ew")
button_browse_report.grid(row=1, column=3)
label_password.grid(row=2, column=0, sticky="e")
entry_password.grid(row=2, column=1, columnspan=2, sticky="ew")
button_search.grid(row=3, column=0, columnspan=4)
label_status.grid(row=4, column=0, columnspan=4)
progress_bar.grid(row=5, column=0, columnspan=4, sticky="ew")

# Configure column weights for resizing
app.columnconfigure(1, weight=1)
app.columnconfigure(2, weight=1)

# Start the main event loop
app.mainloop()

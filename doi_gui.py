import requests
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.utils import get_column_letter

# Function to clean HTML tags from abstract text
def clean_html(raw_html):
    soup = BeautifulSoup(raw_html, "html.parser")
    return soup.get_text()

# Function to fetch metadata for a given DOI from the CrossRef API
def fetch_metadata(doi):
    url = f"https://api.crossref.org/works/{doi}"
    response = requests.get(url, headers={"Accept": "application/json"})

    if response.status_code != 200:
        return {
            "DOI": doi,
            "Title": "Not Found",
            "Authors": "",
            "Emails": "",
            "Publication Date": "",
            "Journal Name": "",
            "Abstract": "",
            "Paper Link": f"https://doi.org/{doi}"
        }

    data = response.json().get('message', {})

    title = data.get('title', ['No Title'])[0]
    abstract = data.get('abstract', '')
    if abstract:
        abstract = clean_html(abstract)
    else:
        abstract = 'No Abstract Available'

    authors = data.get('author', [])
    author_names = []
    emails = []
    for author in authors:
        name = f"{author.get('given', '')} {author.get('family', '')}".strip()
        author_names.append(name)
        email = author.get('email')
        if email:
            emails.append(email)

    pub_date_parts = data.get('published-print', data.get('published-online', {})).get('date-parts', [[]])[0]
    publication_date = "-".join(map(str, pub_date_parts)) if pub_date_parts else ''

    journal = data.get('container-title', [''])[0]

    return {
        "DOI": doi,
        "Title": title,
        "Authors": "; ".join(author_names),
        "Emails": "; ".join(emails) if emails else 'Not Available',
        "Publication Date": publication_date,
        "Journal Name": journal,
        "Abstract": abstract,
        "Paper Link": f"https://doi.org/{doi}"
    }
# Function to fetch metadata for a list of DOIs
def process_dois(dois):
    results = []
    for doi in dois:
        metadata = fetch_metadata(doi)
        results.append(metadata)
    return results
# Function to save metadata to an Excel file and auto-adjust column widths
def save_to_excel(data, filepath):
    df = pd.DataFrame(data)
    df.to_excel(filepath, index=False)

    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        adjusted_width = (max_length + 5)
        ws.column_dimensions[column].width = adjusted_width

    wb.save(filepath)
# Function to show file save dialog and return the chosen path
def browse_save_location():
    filepath = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    return filepath
# Function that gets called when the user clicks "Fetch and Save Metadata"
def start_process():
    doi_text = text_input.get("1.0", tk.END).strip()
    dois = [doi.strip() for doi in doi_text.splitlines() if doi.strip()]

    if not dois:
        messagebox.showwarning("Input Error", "Please enter at least one DOI.")
        return

    results = process_dois(dois)

    save_path = browse_save_location()
    if save_path:
        save_to_excel(results, save_path)
        messagebox.showinfo("Success", f"Metadata saved to {save_path}")
    else:
        messagebox.showwarning("Save Cancelled", "File not saved.")

# GUI Setup
root = tk.Tk()
root.title("DOI Metadata Fetcher")
root.geometry("700x500")

label = tk.Label(root, text="Enter DOIs (one per line):", font=("Arial", 14))
label.pack(pady=10)

text_input = tk.Text(root, height=18, width=80)
text_input.pack(pady=5)

process_button = tk.Button(root, text="Fetch and Save Metadata", command=start_process, font=("Arial", 12))
process_button.pack(pady=20)

root.mainloop()

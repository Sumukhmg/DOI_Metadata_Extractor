# 📄 DOI_Metadata_Extractor

**DOI_Metadata_Extractor** is a Python-based desktop tool that allows users to input one or multiple DOIs (Digital Object Identifiers) and fetch detailed metadata such as:

- ✅ Title  
- ✅ Author names  
- ✅ Email addresses (if available)  
- ✅ Publication date  
- ✅ Journal name  
- ✅ Abstract  
- ✅ Paper link  

The extracted metadata is saved into a clean, well-formatted Excel (.xlsx) file for easy reference or research documentation.

---

## 🖥️ Features

- 🔍 Fetches data via CrossRef API  
- 📋 Batch mode: enter multiple DOIs at once  
- 📁 Exports results to Excel  
- 🧑‍🔬 Useful for researchers, librarians, students  
- 🧠 Clean and simple GUI built with Tkinter  
- 📎 Automatically adjusts Excel column widths for readability  

---

## 🚀 Installation

### 🔧 Prerequisites

Make sure you have **Python 3.7+** installed. You can download it from:  
👉 [https://www.python.org/downloads/](https://www.python.org/downloads/)

### 📦 Install required libraries

Open a terminal or command prompt and run:

```bash
pip install requests pandas openpyxl beautifulsoup4
```

# ▶️ How to Run

### 1. Download or clone the repository:

```bash
git clone https://github.com/Sumukhmg/DOI_Metadata_Extractor.git
cd DOI_Metadata_Extractor
```
### 2. Run the tool:

```bash
python doi_gui.py
```

### 3. Usage:

- Enter one DOI per line in the text area.

- Click "Fetch and Save Metadata".

- Choose where to save the Excel file.


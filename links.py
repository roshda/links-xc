import pandas as pd
import requests
import fitz
import tkinter as tk
from tkinter import filedialog
import os

def count_pdf_pages(url):
    try:
        response = requests.get(url)
        response.raise_for_status()

        with fitz.open("pdf", response.content) as doc:
            title = doc.metadata.get('title', 'No Title Found')
            return len(doc), title
    except Exception as e:
        print(f"Error processing {url}: {e}")
        return "LINK ERROR", 'No Title Found'

def process_excel(file_path):
    df = pd.read_excel(file_path, header=None)
    df[[1, 2]] = df[0].apply(lambda url: pd.Series(count_pdf_pages(url)))
    df.to_excel(file_path, index=False, header=False)

def process_pdf_folder(folder_path):
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    data = []
    for file in pdf_files:
        try:
            with fitz.open(os.path.join(folder_path, file)) as doc:
                data.append([file, len(doc)])
        except Exception as e:
            print(f"Error processing {file}: {e}")
            data.append([file, "ERROR"])

    df = pd.DataFrame(data, columns=['PDF Name', 'Page Count'])
    output_path = os.path.join(folder_path, 'pdf_page_counts.xlsx')
    df.to_excel(output_path, index=False)

def select_excel_file():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])

def select_folder():
    root = tk.Tk()
    root.withdraw()
    return filedialog.askdirectory(title="Select Folder")

def main():
    choice = input("Type '1' to select a folder of PDFs or '2' for an Excel sheet with links: ")
    if choice == '1':
        folder_path = select_folder()
        if folder_path:
            process_pdf_folder(folder_path)
        else:
            print("No folder selected.")
    elif choice == '2':
        excel_file_path = select_excel_file()
        if excel_file_path:
            process_excel(excel_file_path)
        else:
            print("No file selected.")
    else:
        print("Invalid choice.")

if __name__ == "__main__":
    main()

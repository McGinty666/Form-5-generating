!pip install pandas
!pip install openpyxl
!pip install xlrd
!pip install python-docx

import os
import pandas as pd
from docx import Document
from docx.shared import Pt
from copy import deepcopy
from google.colab import files

def read_table_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df

def read_template_document(template_path):
    return Document(template_path)

def add_column_to_word(document, column_data, column_header):
    # Deepcopy the first paragraph of the first cell to preserve formatting
    first_cell_first_paragraph = deepcopy(document.tables[0].cell(0, 0).paragraphs[0])
    
    # Clear existing text in the table
    for row in document.tables[0].rows:
        for cell in row.cells:
            cell.paragraphs[0].clear()
    
    # Set the header
    document.tables[0].cell(0, 0).paragraphs[0].add_run(column_header).font.size = Pt(12)

    # Add the column data
    for i, data in enumerate(column_data, start=1):
        # Add rows to the table dynamically if needed
        while len(document.tables[0].rows) < i + 1:
            document.tables[0].add_row()
        
        document.tables[0].cell(i, 0).paragraphs[0].add_run(str(data)).font.size = Pt(10)

    return document

def process_excel_documents(uploaded_files, template_document):
    excel_documents = {}
    for filename, file_content in uploaded_files.items():
        file_path = filename
        with open(file_path, 'wb') as f:
            f.write(file_content)
        df = read_table_from_excel(file_path)
        for column in df.columns:
            excel_documents[column] = df[column].tolist()

    return template_document, excel_documents

def create_word_documents(template_document, word_data, output_folder):
    output_files = []
    for column_header, column_data in word_data.items():
        new_document = deepcopy(template_document)
        new_document = add_column_to_word(new_document, column_data, column_header)
        output_path = os.path.join(output_folder, f"{column_header}.docx")
        new_document.save(output_path)
        output_files.append(output_path)
        print(f"Word document created at {output_path}")
    return output_files

if __name__ == "__main__":
    uploaded = files.upload()
    print("Please upload the template document (Form5_empty.docx):")
    template_upload = files.upload()
    template_path = list(template_upload.keys())[0]
    template_document = read_template_document(template_path)
    
    output_folder = "/content/output_form5"
    os.makedirs(output_folder, exist_ok=True)

    template_document, word_data = process_excel_documents(uploaded, template_document)
    output_files = create_word_documents(template_document, word_data, output_folder)
    
    for file_path in output_files:
        files.download(file_path)

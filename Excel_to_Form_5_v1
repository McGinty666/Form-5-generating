# -*- coding: utf-8 -*-
"""
Created on Tue Nov 21 21:41:12 2023

@author: robmc
"""


import os
import pandas as pd
from docx import Document
from docx.shared import Pt
from copy import deepcopy

def read_table_from_word(doc_path):
    doc = Document(doc_path)
    tables = doc.tables
    if tables:
        # Assuming you want to extract data from the first table in each document
        table = tables[0]
        data = []
        for row in table.rows:
            row_data = [cell.text for cell in row.cells]
            data.append(row_data)
        return data
    else:
        return None

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
        
        document.tables[0].cell(i, 0).paragraphs[0].add_run(data).font.size = Pt(10)

    return document

def process_word_documents(folder_path, template_path):
    template_document = read_template_document(template_path)
    word_documents = {}
    for filename in os.listdir(folder_path):
        if filename.endswith(".docx"):
            doc_path = os.path.join(folder_path, filename)
            table_data = read_table_from_word(doc_path)
            if table_data:
                word_documents[filename] = table_data

    return template_document, word_documents

def create_word_documents(template_document, word_data, output_folder):
    for filename, column_data in word_data.items():
        new_document = deepcopy(template_document)
        new_document = add_column_to_word(new_document, column_data, filename)
        output_path = os.path.join(output_folder, f"output_{filename}")
        new_document.save(output_path)
        print(f"Word document created at {output_path}")

if __name__ == "__main__":
    folder_path = "E:/Work/ttachments"
    template_path = "E:/Work/ttachments/Form5_empty.docx"
    output_folder = "E:/Work/output_form5"

    template_document, word_data = process_word_documents(folder_path, template_path)
    create_word_documents(template_document, word_data, output_folder)

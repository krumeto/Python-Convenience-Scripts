#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import numpy as np
import time
import openpyxl as xl
import PyPDF2
get_ipython().run_line_magic('matplotlib', 'inline')
from pathlib import Path
import glob
from PyPDF2 import PdfFileReader


# In[ ]:


def pdf_to_test(file_name):
    """
    A function to read a pdf, parse it to text and return a dictionary with values.
    """
    #Opening, reading and parsing a pdf file to string
    pdfFileObj = open(file_name, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    pdf_string = pdfReader.getPage(0).extractText()
    
    #Find the RechnungsNr.
    start_of_RN = pdf_string.find("No.Invoice Date") + len("No.Invoice Date")
    rechnungs_nr = pdf_string[start_of_RN:start_of_RN+7]
    
    #Find the address
    start_of_address = pdf_string.find("Invoice Address") + len("Invoice Address")
    end_of_address = pdf_string.find("Payment Terms:")
    address = pdf_string[start_of_address:end_of_address]
    
    #Liefermonat commenrs
    start_of_contract = pdf_string.find("Company Name / Line of business") + len("Company Name / Line of business")
    end_of_contract = pdf_string.find("Summary of Charges")
    contract = pdf_string[start_of_contract:end_of_contract]
    
    #Nettobetrag - read base charge
    start_of_netto = pdf_string.find("Base Charges") + len("Base Charges")
    end_of_netto = pdf_string.find("Click Charges - Color")
    nettobetrag = pdf_string[start_of_netto:end_of_netto]
    
    pdfFileObj.close()
    
    return pdfFileObj.name, rechnungs_nr, address, contract, nettobetrag


# In[ ]:


def folder_to_df(path):
    """
    Globs through a folder and returns a table after reading the PDF
    """
    summary_df = pd.DataFrame(columns=["file_name", "invoice_nr", "address", "contract", "base_charge"])
    
    for file in Path(path).glob("*.pdf"):
        print(file)
        try: 
            summary_df = summary_df.append({
                "file_name": pdf_to_test(file)[0],
                "invoice_nr": pdf_to_test(file)[1],
                "address": pdf_to_test(file)[2],
                "contract": pdf_to_test(file)[3],
                "base_charge": pdf_to_test(file)[4]}, 
                ignore_index = True)
        except:
            summary_df = summary_df.append({
                "file_name": file.name,
                "invoice_nr": "Could not read malformed PDF file",
                "address": "Could not read malformed PDF file",
                "contract": "Could not read malformed PDF file",
                "base_charge": "Could not read malformed PDF file"}, 
                ignore_index = True)
    return summary_df


# 1. Put the right Windows path below (be careful with / ) :)

# In[ ]:


invoice_path = "C:/Users/arnaudok/OneDrive - Company Name Inc/Python Scripts/Work for specific accounts"
results = folder_to_df(invoice_path)
results.head(10) # this will print the first 10 rows, just so that you know if you are doing well


# 2. Convert to excel

# In[ ]:


results.to_excel("PDF_summary.xlsx", index=False)


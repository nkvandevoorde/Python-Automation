from openpyxl import workbook
import openpyxl
from openpyxl import load_workbook
import unicodedata
import pandas as pd

#Remove special characters for consistent spelling
def remove_accents(input_str):
    if input_str is None:
        return None
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return ''.join([c for c in nfkd_form if not unicodedata.combining(c)])

#Load the Excel files
CNES_sheet = load_workbook(filename='CNES_2024_with_REGIC_labels_p.xlsx').active
CNES_Addresses = load_workbook(filename='cnes_output.xlsx').active

Relatorio_sheet = load_workbook(filename='relatorio-geral_COPY0510.xlsx').active
Relatorio_Addresses = load_workbook(filename='relatorio_output.xlsx').active

#Concatenate number, address, and zip into a single string
def concatenate_address(row):
    number = str(row[0].value) if row[0].value is not None else ''
    address = str(row[1].value) if row[1].value is not None else ''
    zip_code = str(row[2].value) if row[2].value is not None else ''
    full_address = f"{number} {address} {zip_code}".strip()
    return remove_accents(full_address).lower()

#Create a dictionary to store the concatenated addresses from the CNES sheet
cnes_addresses_dict = {}
for row_idx in range(2, CNES_Addresses.max_row + 1):
    cnes_val = CNES_Addresses.cell(row=row_idx, column=1).value 
    cnes_address = concatenate_address(CNES_Addresses[row_idx])
    cnes_addresses_dict[cnes_val] = cnes_address

#Make Address column in CNES sheet for corresponding CNES values
for row_idx in range(2, CNES_sheet.max_row + 1):
    cnes_val = CNES_sheet.cell(row=row_idx, column=1).value 
    if cnes_val in cnes_addresses_dict:
        CNES_sheet.cell(row=row_idx, column=2).value = cnes_addresses_dict[cnes_val]

#Update CNES excel file
CNES_sheet.parent.save('CNES_2024_with_REGIC_labels_p.xlsx')

#Create a dictionary to store the concatenated addresses from the Relatorio sheet
relatorio_addresses_dict = {}
for row_idx in range(2, Relatorio_Addresses.max_row + 1):
    relatorio_val = Relatorio_Addresses.cell(row=row_idx, column=1).value 
    relatorio_address = concatenate_address(Relatorio_Addresses[row_idx])
    relatorio_addresses_dict[relatorio_val] = relatorio_address

#Make Address column in Relatorio sheet for corresponding CNES values
for row_idx in range(2, Relatorio_sheet.max_row + 1):
    relatorio_val = Relatorio_sheet.cell(row=row_idx, column=1).value 
    if relatorio_val in relatorio_addresses_dict:
        Relatorio_sheet.cell(row=row_idx, column=2).value = relatorio_addresses_dict[relatorio_val]

#Update Relatorio excel file
Relatorio_sheet.parent.save('relatorio-general_COPY0510.xlsx')

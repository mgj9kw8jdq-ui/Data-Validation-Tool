# -*- coding: utf-8 -*-
"""
Created on Thu May 29 14:55:29 2025

@author: 20353120
"""

import os
import openpyxl
import pandas as pd
import xlrd
from tqdm import tqdm  # progress bar
# Folder containing the PT reports
folder_path = r'D:\OneDrive - Larsen & Toubro\NDT REPORTS\PT   Report'
# Output data list
data = []
# Get all Excel files
all_files = [f for f in os.listdir(folder_path) if f.endswith(('.xlsx', '.xlsm', '.xls'))]
total_files = len(all_files)
print(f"\n📂 Total files found: {total_files}\n")
# tqdm adds a smooth progress bar
for filename in tqdm(all_files, desc="📊 Processing Reports", unit="file"):
   filepath = os.path.join(folder_path, filename)
   if filename.endswith(('.xlsx', '.xlsm')):
       wb = openpyxl.load_workbook(filepath, data_only=True)
       sheet = wb.worksheets[0]
       report_number = sheet['O3'].value
       report_date = sheet['O4'].value
       for row in range(25, 34):
           line_number = sheet[f'B{row}'].value
           joint_number = sheet[f'E{row}'].value
           material = sheet[f'K{row}'].value
           welder = sheet[f'H{row}'].value
           if line_number or joint_number:
               data.append({
                   'Report Number': report_number,
                   'Report Date': report_date,
                   'Line Number': line_number,
                   'Joint Number': joint_number,
                   'Material': material,
                   'Welder': welder,
                   'File Name': filename
               })
   elif filename.endswith('.xls'):
       wb = xlrd.open_workbook(filepath)
       sheet = wb.sheet_by_index(0)
       report_number = sheet.cell_value(2, 14)  # K8 = row 7, col 10 (0-indexed)
       report_date = sheet.cell_value(3, 14)
       for row in range(24, 33):
           line_number = sheet.cell_value(row, 1)
           joint_number = sheet.cell_value(row, 4)
           material = sheet.cell_value(row, 10)
           welder = sheet.cell_value(row, 7)
           if line_number or joint_number:
               data.append({
                   'Report Number': report_number,
                   'Report Date': report_date,
                   'Line Number': line_number,
                   'Joint Number': joint_number,
                   'Material': material,
                   'Welder': welder,
                   'File Name': filename
               })
# Convert to DataFrame and save
df = pd.DataFrame(data)
output_path = r'D:\OneDrive - Larsen & Toubro\Desktop\PT Summary.xlsx'
df.to_excel(output_path, index=False)
print(f"\n✅ All done. Summary saved to: {output_path}")
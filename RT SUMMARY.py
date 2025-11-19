import os
import openpyxl
import pandas as pd
from tqdm import tqdm

# Folder containing the PT reports
folder_path = r'D:\OneDrive - Larsen & Toubro\NDT REPORTS\RT  Report'
# Output data list
data = []
bad_files = []
# Get all Excel files
all_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.xlsx', '.xlsm'))]
total_files = len(all_files)
print(f"\n📂 Total files found: {total_files}\n")
# tqdm adds a smooth progress bar
for filename in tqdm(all_files, desc="📊 Processing Reports", unit="file"):
   filepath = os.path.join(folder_path, filename)
   try:
       wb = openpyxl.load_workbook(filepath, data_only=True)
       
       sheet = None
       for name in wb.sheetnames:
           sht = wb[name]
           if sht.sheet_state != 'hidden':   # pick first visible sheet
               sheet = sht
               break
           
       if sheet is None:
           print(f"❌ No visible sheets found in file {filename}, skipping.")
           continue
       
       report_number = sheet['P3'].value
       report_date = sheet['P4'].value
       line_number = sheet['B19'].value
       joint_number = sheet['D19'].value
       material = sheet['F19'].value
       welder = sheet['I19'].value
       
       data.append({
                   'Report Number': report_number,
                   'Report Date': report_date,
                   'Line Number': line_number,
                   'Joint Number': joint_number,
                   'Material': material,
                   'Welder': welder,
                   'File Name': filename
               })
       
   except Exception as e:
        print(f"Error reading {filename} : {e}")
        bad_files.append(filename)
   
# Convert to DataFrame and save
df = pd.DataFrame(data)
output_path = r'D:\OneDrive - Larsen & Toubro\Desktop\RT Summary.xlsx'
df.to_excel(output_path, index=False)
print(f"\n✅ All done. Summary saved to: {output_path}")

if bad_files:
    print("\n Files that could not be processed: ")
    for f in bad_files:
        print(f" - {f}")
else:
    print("\n All files processed succesfully without errors.")
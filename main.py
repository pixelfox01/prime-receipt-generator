import pandas as pd
import openpyxl
import os
import sys
import shutil

# === Settings ===
template_file = "template.xlsx"

input_dir = os.path.join(os.getcwd(), "input")
if not os.path.exists(input_dir):
    print(
        "'input' folder not found. Please create a folder named 'input' in the same directory as the .exe."
    )
    sys.exit(1)

# List all .xlsx files in input/ excluding temp files
xlsx_files = [
    f for f in os.listdir(input_dir) if f.endswith(".xlsx") and not f.startswith("~$")
]

if len(xlsx_files) == 0:
    print("No Excel files found in the 'input' folder.")
    sys.exit(1)
elif len(xlsx_files) > 1:
    print("Multiple Excel files found in the 'input' folder. Please keep only one.")
    sys.exit(1)

input_file = os.path.join(input_dir, xlsx_files[0])

# Ensure the output directory exists
output_dir = os.path.join(os.getcwd(), "output")
os.makedirs(output_dir, exist_ok=True)

# Define output path
output_path = os.path.join(output_dir, "filled_receipts.xlsx")

# === Step 1: Load DSR and find where table starts ===
raw_df = pd.read_excel(input_file, header=None)
start_row_idx = raw_df[raw_df.iloc[:, 0] == "Date"].index[0]
df = pd.read_excel(input_file, skiprows=start_row_idx)

date = df.iloc[0]["Date"]
date = date.strftime("%d-%m-%Y")
df = df.iloc[1:]  # Skip header row

# Drop rows after first empty BillNo
billno_col = "BillNo"
if billno_col in df.columns:
    df = df[df[billno_col].notna() & (df[billno_col] != "")]
else:
    raise ValueError(f"'{billno_col}' column not found in input file.")

# === Step 2: Divide data into chunks of 6 (per receipt page) ===
chunks = [df.iloc[i : i + 6] for i in range(0, len(df), 6)]

# === Step 3: Copy template and load workbook ===
shutil.copy(template_file, output_path)
wb = openpyxl.load_workbook(output_path)

# === Step 4: Fill in data and track used sheets ===
used_sheet_names = []

for group_index, chunk in enumerate(chunks):
    if group_index >= len(wb.worksheets):
        print("⚠️ Not enough sheets in template to fit all chunks.")
        break

    sheet = wb.worksheets[group_index]
    used_sheet_names.append(sheet.title)

    for i, (_, row) in enumerate(chunk.iterrows(), start=1):
        store = row["Party"]
        bill = row["BillNo"]
        amount = row["Gross"]

        for row_cells in sheet.iter_rows():
            for cell in row_cells:
                if isinstance(cell.value, str):
                    cell.value = (
                        cell.value.replace(f"{{{{DATE{i}}}}}", str(date))
                        .replace(f"{{{{PARTY{i}}}}}", str(store))
                        .replace(f"{{{{BILL NO{i}}}}}", str(bill))
                        .replace(f"{{{{AMOUNT{i}}}}}", f"{amount:.2f}")
                    )

# === Step 5: Delete unused sheets ===
for sheet in wb.worksheets[len(used_sheet_names) :]:
    print(f"🗑️ Deleting unused sheet: {sheet.title}")
    wb.remove(sheet)

wb.save(output_path)

print(f"File generated: {output_path}")

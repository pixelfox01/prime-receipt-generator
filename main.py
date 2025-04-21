import pandas as pd
import openpyxl
import os
import sys

from pathlib import Path

def process_excel(input_path):
    # Load data and find where table starts
    raw_df = pd.read_excel(input_path, header=None)
    start_row_idx = raw_df[raw_df.iloc[:, 0] == "Date"].index[0]
    df = pd.read_excel(input_path, skiprows=start_row_idx)

    date = df.iloc[0]["Date"]
    df_skipped = df.iloc[1:]

    # Load template workbook and sheet
    template_path = os.path.join(os.getcwd(), "template.xlsx")
    if not os.path.exists(template_path):
        raise FileNotFoundError("template.xlsx not found in the current directory.")

    wb = openpyxl.load_workbook(template_path)
    template_sheet = wb.active

    stop = False
    chunk_index = 1

    for i in range(0, len(df_skipped), 6):
        chunk = df_skipped.iloc[i : i + 6]

        new_sheet = wb.copy_worksheet(template_sheet)
        new_sheet.title = f"Group {chunk_index}"

        for j, (_, row) in enumerate(chunk.iterrows(), start=1):
            bill_no, party, net_amt = row["BillNo"], row["Party"], row["NetAmt"]

            if pd.isna(bill_no) or bill_no == "":
                stop = True
                break

            for row_cells in new_sheet.iter_rows():
                for cell in row_cells:
                    if isinstance(cell.value, str):
                        cell.value = cell.value.replace(f"{{{{DATE{j}}}}}", str(date))
                        cell.value = cell.value.replace(f"{{{{BILL NO{j}}}}}", str(bill_no))
                        cell.value = cell.value.replace(f"{{{{PARTY{j}}}}}", str(party))
                        cell.value = cell.value.replace(f"{{{{AMOUNT{j}}}}}", f"{net_amt:.2f}")

        for row_cells in new_sheet.iter_rows():
            for cell in row_cells:
                if isinstance(cell.value, str) and "report_date" in cell.value:
                    cell.value = cell.value.replace("report_date", str(date))

        chunk_index += 1
        if stop:
            break

    del wb[template_sheet.title]

    # Ensure the output directory exists
    output_dir = os.path.join(os.getcwd(), "output")
    os.makedirs(output_dir, exist_ok=True)

    # Define output path
    output_path = os.path.join(output_dir, "filled_all_chunks.xlsx")
    wb.save(output_path)

    print(f"File generated: {output_path}")


if __name__ == "__main__":
    input_dir = os.path.join(os.getcwd(), "input")
    if not os.path.exists(input_dir):
        print("'input' folder not found. Please create a folder named 'input' in the same directory as the .exe.")
        sys.exit(1)

    # List all .xlsx files in input/ excluding temp files
    xlsx_files = [
        f for f in os.listdir(input_dir)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]

    if len(xlsx_files) == 0:
        print("No Excel files found in the 'input' folder.")
        sys.exit(1)
    elif len(xlsx_files) > 1:
        print("Multiple Excel files found in the 'input' folder. Please keep only one.")
        sys.exit(1)

    input_file_path = os.path.join(input_dir, xlsx_files[0])

    try:
        process_excel(input_file_path)
    except Exception as e:
        print(f"Error processing file: {e}")

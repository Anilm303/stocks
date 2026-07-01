import pandas as pd
import os
import glob

def convert_xlsx_to_csv(input_dir):
    # Find all .xlsx files in the directory
    xlsx_files = glob.glob(os.path.join(input_dir, "*.xlsx"))

    if not xlsx_files:
        print(f"No .xlsx files found in {input_dir}")
        return

    for file_path in xlsx_files:
        # Skip temporary Excel files
        if os.path.basename(file_path).startswith("~$"):
            continue

        print(f"Converting: {os.path.basename(file_path)}")
        try:
            # Read the Excel file
            # Note: By default, this reads the first sheet.
            # If there are multiple sheets, we might need to handle them.
            df = pd.read_excel(file_path)

            # Construct the CSV file path
            csv_file_path = os.path.splitext(file_path)[0] + ".csv"

            # Save as CSV
            df.to_csv(csv_file_path, index=False)
            print(f"Successfully converted to: {os.path.basename(csv_file_path)}")
        except Exception as e:
            print(f"Failed to convert {os.path.basename(file_path)}: {e}")

if __name__ == "__main__":
    target_directory = r"C:/Users/Lenovo/Desktop/stocks/Excel_Files"
    convert_xlsx_to_csv(target_directory)

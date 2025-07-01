import pandas as pd
import os

def split_by_tea_estate(input_file, output_file):
    # Read the Excel file
    df = pd.read_excel(input_file)
    header = df.columns.tolist()
    
    # Handle blanks in Column C (index 2)
    df[df.columns[2]] = df[df.columns[2]].fillna("Blank")

    # Create a new ExcelWriter
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for estate, group in df.groupby(df.columns[2]):
            group.to_excel(writer, sheet_name=str(estate)[:31], index=False)

    print(f"✅ Done: Created '{output_file}' with one sheet per tea estate.")

# Example usage:
if __name__ == "__main__":
    input_path = "tea_data.xlsx"       # Put your Excel file name here
    output_path = "tea_estates_grouped.xlsx"
    split_by_tea_estate(input_path, output_path)

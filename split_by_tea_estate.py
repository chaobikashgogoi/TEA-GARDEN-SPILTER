import pandas as pd
import os
import re

def clean_sheet_name(name):
    """
    Sanitize sheet name by removing or replacing invalid characters
    and ensuring it meets Excel's requirements.
    """
    # Convert to string and truncate to 31 characters
    name = str(name)[:31]
    # Remove invalid characters: / \ * ? : [ ]
    name = re.sub(r'[\/\\*?:\[\]]', '_', name)
    # If name is empty after cleaning, provide a default name
    return name if name.strip() else 'Sheet'

def split_by_tea_estate(input_file, output_file):
    # Read the Excel file
    df = pd.read_excel(input_file)
    header = df.columns.tolist()
    
    # Handle blanks in Column B (index 1)
    df[df.columns[1]] = df[df.columns[1]].fillna("Blank")

    # Create a new ExcelWriter
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for estate, group in df.groupby(df.columns[1]):
            # Use cleaned sheet name
            sheet_name = clean_sheet_name(estate)
            group.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"✅ Done: Created '{output_file}' with one sheet per value in Column B.")

# Example usage:
if __name__ == "__main__":
    input_path = "tea_data.xlsx"       # Put your Excel file name here
    output_path = "tea_estates_grouped.xlsx"
    split_by_tea_estate(input_path, output_path)

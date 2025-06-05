
#This script extracts tables from a PDF file using the tabula-py library, cleans and splits the data, and saves it to an Excel file.

import pandas as pd
from tabula import read_pdf
import re
from typing import Dict
import unicodedata

#aeris, cimed 

# Specify the range of pages to extract tables from

### corsan 15 to 16 ### brisanet 13 to 15 ### matix 8 to 10 ### vibra 3 to 4 ### vli 14 to 15 ###
start_page = 8  # Change this to the starting page number 
end_page = 10  # Change this to the ending page number

# Extract the PDF file name without extension
#pdf_name = "vibraITR"  # Replace with dynamic extraction if needed
pdf_name = "matrixITR"  # Replace with dynamic extraction if needed
#pdf_name = "vibraITR"  # Replace with dynamic extraction if needed
#pdf_name = "vibraITR"  # Replace with dynamic extraction if needed
#pdf_name = "vibraITR"  # Replace with dynamic extraction if needed

import re
from typing import Union

def clean_and_split_cell(cell: Union[str, float, int]) -> Union[str, float, int]:
    """
    Cleans and splits concatenated or improperly formatted content in a cell.

    ----------------------------------------
    REGEX DICTIONARY (used in this function)
    ----------------------------------------
    \d       → any digit (0–9)
    \d+      → one or more digits (e.g., 1, 123, 9999)
    \d{3}    → exactly three digits
    \s       → any whitespace character (space, tab, newline)
    \s+      → one or more whitespace characters
    \b       → word boundary (e.g., end of a word/number)
    \.       → literal dot (escaped, since dot means "any character" in regex)
    (?:...)  → non-capturing group (used for grouping without backreference)
    (?=...)  → lookahead (matches if followed by pattern but doesn't consume it)
    (?<=...) → lookbehind (matches if preceded by pattern but doesn't consume it)
    (...)    → capturing group (captures content for substitution/backreference)
    |        → OR (used to match one pattern or another)

    Returns:
        Cleaned and transformed cell content.
    """
    
    if not isinstance(cell, str):
        return cell

    # Normalize to NFC and remove invisible characters
    cell = unicodedata.normalize("NFKC", cell)
    cell = re.sub(r'[\u200b\u200e\u202f\xa0]', '', cell)  # remove invisibles

    cell = cell.strip()

    # Handle: "70-" or "70 -" → separate number and trailing hyphen
    cell = re.sub(r'(\d+)\s*-\s*$', r'\1;-', cell)

    # Handle: "-47  -" → two separate values "-47" and "-"
    cell = re.sub(r'(-?\d+)\s+-\b', r'\1;-', cell)

    # Handle: malformed leading dash space dash → "- ;num"
    cell = re.sub(r'-\s+(-?\d+)', r'-;\1', cell)

    # Handle: two numbers separated by space → "num;num"
    cell = re.sub(
        r'(?<![\d/])(\d{1,3}(?:\.\d{3})*|\d+)\s+(\d{1,3}(?:\.\d{3})*|\d+)(?![\d/])',
        r'\1;\2',
        cell
    )
    
    # Handle: "PassivoNota" → "Passivo Nota"
    cell = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', cell)


    # Handle: "Receitadevendas" → "Receita de vendas" by adding space
    cell = re.sub(r'([a-z])([A-Z])', r'\1 \2', cell)

    # Handle: "Receita2023" → "Receita 2023"
    cell = re.sub(r'(?<=[a-zA-Z])(?=\d)', ' ', cell)

    # Handle: "172.272162.947" → "172.272 162.947"
    cell = re.sub(r'(?<=\d)(?=\d{3}\b)', ' ', cell)

    # Remove unnecessary space between letters (e.g., "Receita de" stays the same)
    #cell = re.sub(r'(?<=[a-zA-Z]) (?=[a-zA-Z])', '', cell)

    # Handle: "31/12/2 022" → "31/12/2022"
    cell = re.sub(r'(\d{2})\s*/\s*(\d{2})\s*/\s*(\d{1})\s*(\d{3})', r'\1/\2/\3\4', cell)

    # Handle: "31 / 12 / 2022" → "31/12/2022"
    cell = re.sub(r'(\d{2})\s*/\s*(\d{2})\s*/\s*(\d{4})', r'\1/\2/\3', cell)

    # Handle: "2 024" → "2024"
    cell = re.sub(r'(\d)\s+(\d{3})', r'\1\2', cell)

    # Handle: "2024 2023" → "2024;2023"
    cell = re.sub(r'(\d{4})\s+(\d{4})', r'\1;\2', cell)

    # Handle: "8631.224" → "863;1.224"
    cell = re.sub(r'(\d{3})(\d\.\d{3})', r'\1;\2', cell)

    # Handle: "321.323525" → "321.323;525"
    cell = re.sub(r'(\d{1,3}(?:\.\d{3})+)(\d{3})\b', r'\1;\2', cell)

    # Handle: "10.4806.666" → "10.480;6.666"
    cell = re.sub(r'(\d+\.\d{3})(\d{3}\b)', r'\1;\2', cell)

    # Handle: "10.4 806.666" → "10.4;806.666"
    cell = re.sub(r'(\d+\.\d)\s+(\d{3}\.\d{3})', r'\1;\2', cell)

    # Handle: merged groups of formatted numbers
    cell = re.sub(r'(\d{1,3}(?:\.\d{3})+)(\d{1,3}(?:\.\d{3})+)', r'\1;\2', cell)

    # Handle: "(164.031)  (154.586)" → "(-164.031);(-154.586)"
    cell = re.sub(r'\((\d+\.\d{3})\)\s+\((\d+\.\d{3})\)', r'(\1);(\2)', cell)

    # Handle: "(868)  (780)" → "(868);(780)"
    cell = re.sub(r'\((\d+)\)\s+\((\d+)\)', r'(\1);(\2)', cell)

    # Handle: "(113) 519" → "(113);519"
    cell = re.sub(r'\((\d+)\)\s+(\d+)', r'(\1);\2', cell)

    # Handle: "511  (791)" → "511;(791)"
    cell = re.sub(r'(\d+)\s+\((\d+)\)', r'\1;(\2)', cell)

    # Handle: "511  (791.685)" → "511;(791.685)"
    cell = re.sub(r'(\d+)\s+\((\d+\.\d{3})\)', r'\1;(\2)', cell)

    # Handle: "(113.465) 519" → "(113.465);519"
    cell = re.sub(r'\((\d+\.\d{3})\)\s+(\d+)', r'(\1);\2', cell)

    # Handle: "(113)(154.586)" → "(113);(154.586)"
    cell = re.sub(r'\((\d+)\)\((\d+\.\d{3})\)', r'(\1);(\2)', cell)

    # Handle: "(2.099)(1.459)" → "(2.099);(1.459)"
    cell = re.sub(r'\((\d+\.\d{3})\)\((\d+\.\d{3})\)', r'(\1);(\2)', cell)

    # Handle: "(2.099)(449)" → "(2.099);(449)"
    cell = re.sub(r'\((\d+\.\d{3})\)\((\d+)\)', r'(\1);(\2)', cell)

    # Handle: "(95)(65)" → "95;(65)"
    cell = re.sub(r'\((\d+)\)\((\d+)\)', r'\1;(\2)', cell)

    # Handle: "196(4.051)" → "196;(4.051)"
    cell = re.sub(r'(\d+)\((\d+\.\d{3})\)', r'\1;(\2)', cell)

    # Convert: "(123.456)" → "-123.456" and "(123)" → "-123"
    cell = re.sub(r'\((\d+(?:\.\d{3})?)\)', r'-\1', cell)

    # Ex: "Fornecedores 13" ou "Nota5" → "Fornecedores;13"
    cell = re.sub(r'([a-zA-Z]+)\s*(\d{1,3})\b', r'\1;\2', cell)

    cell = re.sub(r'(-?\d+)\s+-\b', r'\1;-', cell)

    return cell.strip()

def preprocess_table(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        df[col] = df[col].apply(clean_and_split_cell)
    return df

# Function to manually split concatenated columns if automatic extraction fails
def manually_split_columns(df: pd.DataFrame) -> pd.DataFrame:
    for col in df.columns:
        if df[col].dtype == 'object':
            df[col] = df[col].apply(
                lambda x: re.split(r'(?<=\d{3})(?=\d{3})', x)
                if isinstance(x, str) and re.search(r'\d{3}\d{3}', x) else x
            )
            df[col] = df[col].apply(lambda x: ";".join(x) if isinstance(x, list) else x)
    return df

# Function to split semicolon-separated or space-separated values into separate columns
def split_semicolon_values(df):
    new_columns = []
    for col in df.columns:
        if df[col].dtype == 'object' and df[col].str.contains(";").any():

            # Split semicolon-separated values into separate columns
            split_data = df[col].str.split(";", expand=True)
            # Rename new columns to avoid conflicts
            split_data.columns = [f"{col}_split_{i+1}" for i in range(split_data.shape[1])]
            new_columns.append(split_data)
        else:
            # Keep the original column if no splitting is needed
            new_columns.append(df[[col]])
    
    # Combine all columns back into a single DataFrame
    df = pd.concat(new_columns, axis=1)

    return df

# Extract tables using tabula-py with stream mode for better cell separation
tables = read_pdf(
    f"C:\\Users\\jaotr\\OneDrive\\Documentos\\itau_am\\py_table_scammer\\{pdf_name}.pdf",
    pages=f"{start_page}-{end_page}",
    multiple_tables=True,
    stream=True,  # Use stream mode to handle concatenated cells better
    pandas_options={"header": None}  # Avoid misinterpreting the first row as headers
)

# Initialize a list to store DataFrames
dataframes = []

# Save tables
with pd.ExcelWriter(f"{pdf_name}_pages_{start_page}_to_{end_page}.xlsx") as writer:
    for i, table in enumerate(tables, start=1):
        # Preprocess the table to clean and split data
        table = preprocess_table(table)
        # Apply manual column splitting as a fallback
        table = manually_split_columns(table)
        # Split semicolon-separated or space-separated values as the last step
        table = split_semicolon_values(table)

        # Drop rows and columns that are completely empty
        table = table.dropna(how='all', axis=0)  # Drop empty rows
        table = table.dropna(how='all', axis=1)  # Drop empty columns

        # Skip writing if the table is empty
        if table.empty:
            print(f"Table {i} from pages {start_page}-{end_page} is empty and will be skipped.")
            continue

        # Append the table to the list of DataFrames
        dataframes.append(table)

        # Save as Excel without row and column indices
        sheet_name = f"Pages{start_page}_to_{end_page}_Table{i}"
        table.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
        print(f"Table {i} from pages {start_page}-{end_page} saved to Excel without indices.")


# Combine all DataFrames into a single DataFrame for future modifications
if dataframes:
    combined_dataframe = pd.concat(dataframes, ignore_index=True)
    print("All tables combined into a single DataFrame for future modifications.")
else:
    print("No tables were extracted or combined.")





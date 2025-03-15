import pandas as pd


def clean_excel(
    input_file,
    output_file,
    sheet_name,
    column_name,
    codes_file,
    codes_sheet,
    codes_column,
):
    """
    Removes rows from an Excel sheet based on a list of codes read from another Excel file.

    Parameters:
        input_file (str): Path to the input Excel file.
        output_file (str): Path to save the cleaned Excel file.
        sheet_name (str): The sheet in the input file to process.
        column_name (str): The name of the column to check for removal.
        codes_file (str): Path to the Excel file containing codes to remove.
        codes_sheet (str): The sheet in the codes file.
        codes_column (str): The column name containing the removal codes.
    """

    # Load the input Excel file and the sheet to be cleaned
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    # Load the Excel file that contains the codes to remove
    df_codes = pd.read_excel(codes_file, sheet_name=codes_sheet)

    # Trim spaces from column names
    df.columns = df.columns.str.strip()
    df_codes.columns = df_codes.columns.str.strip()

    # Convert columns to string to ensure correct matching
    df[column_name] = df[column_name].astype(str).str.strip()
    df_codes[codes_column] = df_codes[codes_column].astype(str).str.strip()

    # Extract codes to remove from the codes file
    codes_to_remove = df_codes[codes_column].dropna().tolist()
    print(f"Codes to remove: {codes_to_remove}")  # Debugging step

    if not codes_to_remove:
        print("Warning: No codes found to remove. Check the codes file!")
        return

    # Remove rows where the column matches codes_to_remove
    df_cleaned = df[~df[column_name].isin(codes_to_remove)]

    # Save the cleaned data
    df_cleaned.to_excel(output_file, sheet_name=sheet_name, index=False)

    print(f"Cleaned file saved as {output_file}")


# Example Usage
input_file = (
    r"D:\Back up 29.11.2023\Paradise Charity\March 2025\LatePayments-2025-03-15.xlsx"
)
output_file = (
    r"D:\Back up 29.11.2023\Paradise Charity\March 2025\LatePaymentsCleaned.xlsx"
)
sheet_name = "Sheet1"
column_name = "SponsorFileNumber"

codes_file = r"D:\Back up 29.11.2023\Paradise Charity\March 2025\CodesToRemove.xlsx"
codes_sheet = "Sheet1"  # Adjust if needed
codes_column = "SponsorID"  # The column in the codes file containing codes

clean_excel(
    input_file,
    output_file,
    sheet_name,
    column_name,
    codes_file,
    codes_sheet,
    codes_column,
)

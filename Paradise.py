import pandas as pd


def clean_excel(input_file, output_file, column_name, codes_to_remove):
    """
    Removes rows from an Excel sheet based on a list of codes.

    Parameters:
        input_file (str): Path to the input Excel file.
        output_file (str): Path to save the cleaned Excel file.
        column_name (str): The name of the column to search for the codes.
        codes_to_remove (list): List of codes to be removed from the specified column.
    """
    # Load the Excel file
    df = pd.read_excel(input_file, sheet_name=sheet_name)
    df[column_name] = df[column_name].astype(str)
    # Filter out rows where the column contains any of the codes
    df_cleaned = df[~df[column_name].isin(codes_to_remove)]

    # Save the cleaned data back to a new Excel file
    df_cleaned.to_excel(output_file, index=False)

    print(f"Cleaned file saved as {output_file}")


# Example usage
input_file = r"D:\Back up 29.11.2023\Paradise Charity\March 2025\LatePayments-2025-03-15.xlsx"  # Change this to your file path
output_file = r"D:\Back up 29.11.2023\Paradise Charity\March 2025\LatePaymentsCleaned.xlsx"  # Change this to your desired output file
sheet_name = "Sheet1"  # Change this to the specific sheet name
column_name = "SponsorFileNumber"  # Change this to the actual column name
codes_to_remove = ["900002"]  # List of codes to remove

clean_excel(input_file, output_file, column_name, codes_to_remove)

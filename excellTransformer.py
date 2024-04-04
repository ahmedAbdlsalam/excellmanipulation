import pandas as pd
import os

def process_excel(input_file, column_name):
    # Get the base filename without extension
    base_filename = os.path.splitext(os.path.basename(input_file))[0]

    # Create a directory with the base filename if it doesn't exist
    output_dir = f"{base_filename}_processed"
    os.makedirs(output_dir, exist_ok=True)

    # Read the Excel file
    xls = pd.ExcelFile(input_file)

    # Iterate through each sheet in the Excel file
    for sheet_name in xls.sheet_names:
        # Read the current sheet
        df = pd.read_excel(input_file, sheet_name=sheet_name)

        # Get unique values from the specified column
        unique_values = df[column_name].unique()

        # Iterate through each unique value
        for value in unique_values:
            # Filter rows based on the unique value
            filtered_rows = df[df[column_name] == value]

            # Create a directory for the current unique value
            value_dir = os.path.join(output_dir, str(value))
            os.makedirs(value_dir, exist_ok=True)

            # Construct the output file path
            output_file = os.path.join(value_dir, f"{sheet_name}.xlsx")

            # Write the filtered rows to the output Excel file
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Write the first row of the original sheet
                df.head(0).to_excel(writer, index=False, header=False)
                # Write the filtered rows
                filtered_rows.to_excel(writer, index=False, header=False)

# Example usage
input_file = "sheet.xlsx"  # Update with your Excel file path
column_name = "AppName"  # Update with the name of the column you want to use for splitting
process_excel(input_file, column_name)

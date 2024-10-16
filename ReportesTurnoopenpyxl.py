import os
import pandas as pd
import openpyxl
import logging
import warnings

logging.basicConfig(level=logging.INFO)

warnings.filterwarnings("ignore", category=UserWarning,
                        message="Data validation extension is not supported and will be removed")
warnings.filterwarnings("ignore", category=UserWarning,
                        message="Conditional formatting extension is not supported and will be removed")


def process_file(file_path):
    """
    Processes an Excel file, extracting data from all sheets and returning it as a DataFrame.
    Skips hidden or 'Hoja1' sheets and grabs all data from columns A to K (ignores data beyond column K).
    """
    logging.info(f"Processing file: {file_path}")

    try:
        # Open the Excel workbook
        workbook = openpyxl.load_workbook(file_path, data_only=True)
        all_data = []

        # Loop through all sheets in the file, skip 'Hoja1' or hidden sheets
        for sheet_name in workbook.sheetnames:
            if sheet_name == 'Hoja1' or workbook[sheet_name].sheet_state == 'hidden':
                logging.info(f"Skipping sheet: {sheet_name}")
                continue

            logging.info(f"Processing sheet: {sheet_name}")
            sheet = workbook[sheet_name]
            data = sheet.values

            # Convert the sheet data into a list to manipulate
            data = list(data)

            # Ensure at least 5 rows exist (since headers are in row 4)
            if len(data) < 5:
                logging.warning(f"Sheet {sheet_name} in {file_path} has less than 5 rows, skipping")
                continue

            # Set the 4th row as the header, and everything from row 5 onwards as data
            headers = data[4][:11]  # Grab only up to column K (11 columns)
            rows = [row[:11] for row in data[5:]]  # Append all rows as they appear

            if not rows:
                logging.warning(f"No valid rows found in sheet {sheet_name} in {file_path}, skipping")
                continue

            # Convert all values to string to avoid type issues
            rows = [[str(cell) if cell is not None else '' for cell in row] for row in rows]

            # Create DataFrame with the data restricted to columns A to K
            df = pd.DataFrame(rows, columns=headers)

            # Apply the fix for filtering unnamed columns safely
            try:
                df = df.loc[:, df.columns.str.contains('^Unnamed') == False]
            except Exception as e:
                logging.error(f"Error filtering unnamed columns: {e}")
                continue

            # Remove extra spaces from column names
            df.columns = df.columns.str.strip()

            logging.info(f"Columns in sheet {sheet_name}: {df.columns.tolist()}")

            # Add filename and sheet name for tracking
            df['File Name'] = os.path.basename(file_path)
            df['Sheet Name'] = sheet_name

            all_data.append(df)

        # Return the combined DataFrame
        if not all_data:
            logging.warning(f"No valid data found in {file_path}, returning None")
            return None

        return pd.concat(all_data, ignore_index=True)

    except Exception as e:
        logging.error(f"Error processing file {file_path}: {str(e)}", exc_info=True)
        return None


def process_year_folder(input_folder, output_file_path):
    """
    Processes all files in a folder, merges their sheets, and saves to the output file.
    Focuses on columns A to K in all sheets, skipping 'Hoja1' or hidden sheets.
    """
    files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx')]

    merged_data = []

    for i, file_name in enumerate(files):
        file_path = os.path.join(input_folder, file_name)
        logging.info(f"Processing file {i + 1}/{len(files)}: {file_name}")
        df = process_file(file_path)

        if df is not None:
            merged_data.append(df)

    if not merged_data:
        logging.error(f"No valid data processed from the folder {input_folder}.")
        return

    # Merge all data into a single DataFrame
    merged_df = pd.concat(merged_data, ignore_index=True)

    logging.info(f"Final merged columns: {merged_df.columns.tolist()}")

    # Add new columns 'Ingeniero de Calidad', 'Status', 'Método 7M', 'Fecha'
    merged_df['Ingeniero de Calidad'] = ''
    merged_df['Status'] = ''
    merged_df['Método 7M'] = ''
    merged_df['Fecha'] = ''

    # Save the merged data to an .xlsm file, leaving 5 rows for headers
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        merged_df.to_excel(writer, index=False, startrow=5)
        logging.info(f"Saved merged data to {output_file_path}")


if __name__ == "__main__":
    input_folder = r"C:\Ingenieria en Sis Comp\Practicas\Kayser\ReporteTurnos\2023"
    output_file = r"C:\Ingenieria en Sis Comp\Practicas\Kayser\ReporteTurnos\RP2023.xlsx"
    process_year_folder(input_folder, output_file)

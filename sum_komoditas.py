"""
Komoditas Data Summarizer

This script processes and summarizes komoditas (commodity) data from Excel files
for a specific kabupaten (district) in Indonesia.

@author: Fajrian Aidil Pratama
@email: fajrianaidilp@gmail.com
@created: 2024-08-15 23:47:16
@last_modified: 2024-08-15 23:47:16
@description: Processes Excel files containing komoditas data and generates summary reports
"""
import pandas as pd
import glob
import os

def get_input_with_default(prompt: str, default: any, input_type: type = str) -> any:
    """
    Get user input with a default value and type checking.
    
    :param prompt: The prompt to display to the user
    :param default: The default value if user input is empty
    :param input_type: The expected type of the input (default is str)
    :return: The user input or default value, converted to the specified type
    """
    while True:
        user_input = input(f"{prompt} (default: {default}): ")
        if user_input == "":
            return default
        try:
            return input_type(user_input)
        except ValueError:
            print(f"Invalid input. Please enter a valid {input_type.__name__}.")

def process_files(directory, kab_code, table_prefix):
    """
    Process Excel files in the given directory for a specific kabupaten and table prefix.
    
    :param directory: Path to the directory containing Excel files
    :param kab_code: Kabupaten code to filter data
    :param table_prefix: Prefix for the Excel sheet names
    :return: Tuple of summary_kab (DataFrame) and result_kec (DataFrame)
    """
   # Filter files based on the table_prefix
    files = glob.glob(os.path.join(directory, f'*{table_prefix}*.xlsx'))
    
    if not files:
        print(f"No Excel files found with prefix '{table_prefix}' in directory: {directory}")
        return None, None

    result_kab = pd.DataFrame()
    result_kec = pd.DataFrame()
    
    # These columns definition is used for table 4.54
    # Update these lists for different tables
    columns_to_sum = ['n_rtup_tunggal', 'n_rtup_campuran', 'n_rtup_tumpang_sari', 
                      'n_rtup_asosiasi_antar_semusim', 'n_rtup_jumlah', 'n_rtup']
    
    for file in files:
        print(f"Processing file: {file}")
        try:
            # Process komoditas_kab sheet
            df_kab = pd.read_excel(file, sheet_name=f'{table_prefix}_komoditas_kab')
            district_data_kab = df_kab[df_kab['kab'] == kab_code]
            result_kab = pd.concat([result_kab, district_data_kab], ignore_index=True)
            
            # Process komoditas_kec sheet
            df_kec = pd.read_excel(file, sheet_name=f'{table_prefix}_komoditas_kec')
            district_data_kec = df_kec[df_kec['kab'] == kab_code]
            
            # Sum columns for each kecamatan in this file
            summed_kec = district_data_kec.groupby('kec')[columns_to_sum].sum().reset_index()
            summed_kec['file'] = os.path.basename(file)
            
            result_kec = pd.concat([result_kec, summed_kec], ignore_index=True)
        except Exception as e:
            print(f"Error processing file {file}: {str(e)}")

    # Process kabupaten data
    if result_kab.empty:
        print(f"No data found for id_kab {kab_code} in the {table_prefix}_komoditas_kab sheets.")
        summary_kab = None
    else:
        # Select numeric columns, excluding 'id_komoditas' and 'ID'
        numeric_columns = result_kab.select_dtypes(include=['float64', 'int64']).columns
        row_to_sum = [col for col in numeric_columns if col not in ['id_komoditas', 'ID']]
        
        # Sum the selected numeric columns
        sum_row_kab = result_kab[row_to_sum].sum()
        summary_kab = pd.DataFrame(sum_row_kab).T
        
        # Add non-numeric columns
        for col in ['id_prov', 'id_kab']:
            if col in result_kab.columns:
                new_col = col.replace('id_', '')
                summary_kab[new_col] = result_kab[col].iloc[0]
            else:
                print(f"Column {col} not found in the data.")

    # Process kecamatan data
    if result_kec.empty:
        print(f"No data found for id_kab {kab_code} in the {table_prefix}_komoditas_kec sheets.")
    else:
        # Sum the data for each kecamatan across all files
        result_kec = result_kec.groupby('kec')[columns_to_sum].sum().reset_index()
        
        # Add id columns
        for col in ['id_prov', 'id_kab', 'id_kec']:
            if col in result_kec.columns:
                new_col = col.replace('id_', '')
                result_kec[new_col] = result_kec[col].astype(str)
            else:
                print(f"Column {col} not found in the kecamatan data.")
        
        # Calculate SUM for each kecamatan
        result_kec['SUM'] = result_kec[columns_to_sum].sum(axis=1)

    return summary_kab, result_kec

# Main execution
if __name__ == "__main__":
    # Get user input
    komoditas = input("Enter the name of the komoditas (e.g., jeruk): ")
    kab_code = int(get_input_with_default("Enter the kabupaten code (e.g., 7205): ", 7205, int))
    table_prefix = get_input_with_default("Enter the table number prefix (e.g., 4_54): ", "4_54", str)

    # Set base directory
    # Change this to your specific directory path
    base_directory = ***REMOVED***

    # Construct the full directory path
    directory = os.path.join(base_directory, komoditas)

    # Process the files
    summary_kab, result_kec = process_files(directory, kab_code, table_prefix)

    # Get kabupaten name from the data
    if summary_kab is not None and 'kab' in summary_kab.columns:
        kab_name = str(summary_kab['kab'].iloc[0])
        # Remove the [kab_code] part if it exists
        kab_name = kab_name.split(']')[-1].strip()
    else:
        kab_name = "unknown"

    # Prepare the output file name
    output_file = f"summary_komoditas_{komoditas.lower()}_{kab_name.lower()}.xlsx"

    # Create a Pandas Excel writer using openpyxl as the engine
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        if summary_kab is not None:
            summary_kab.to_excel(writer, sheet_name='Kabupaten', index=False)
            print(f"Saved Kabupaten data to {output_file}")
        
        if not result_kec.empty:
            result_kec.to_excel(writer, sheet_name='Kecamatan', index=False)
            print(f"Saved Kecamatan data to {output_file}")

    print(f"Data has been saved to {output_file}")
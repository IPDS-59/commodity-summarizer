# Komoditas Data Summarizer

This Python script processes and summarizes komoditas (commodity) data from Excel files for a specific kabupaten (district) in Indonesia. It's designed to work with data from Badan Pusat Statistik (BPS) publications.

## Features

- Processes multiple Excel files in a specified directory
- Filters files based on the specified table prefix
- Summarizes data at both kabupaten and kecamatan (sub-district) levels
- Allows customization of input parameters (komoditas, kabupaten code, table prefix)
- Outputs results to a single Excel file with separate sheets for kabupaten and kecamatan data

## Prerequisites

- Python 3.6 or higher
- pandas
- openpyxl

## Installation

1. Clone this repository or download the `sum_komoditas.py` file.
2. Install the required packages:

```
pip install pandas openpyxl
```

## Usage

1. Place your Excel files in a directory structure like this:
   ```
   /path/to/base/directory/
   └── komoditas_name/
       ├── data_4_54_file1.xlsx
       ├── 4_54_file2.xlsx
       └── ...
   ```

   Note: Ensure your Excel files include the table prefix (e.g., '4_54') in their filenames.

2. Run the script:
   ```
   python sum_komoditas.py
   ```

3. Follow the prompts to enter:
   - Name of the komoditas (e.g., jeruk)
   - Kabupaten code (e.g., 1404)
   - Table number prefix (e.g., 4_54)

4. The script will process the files that match the specified table prefix and generate an output Excel file named `summary_komoditas_{komoditas}_{kabupaten}.xlsx` in the same directory as the script.

## Customization

- Update the `base_directory` variable in the script to match your file structure.
- Modify the `columns_to_sum` list if working with a different table structure.

## File Naming Convention

To ensure proper processing, your Excel files should include the table prefix in their names. For example:
- `data_4_54_2023.xlsx`
- `4_54_komoditas_report.xlsx`

The script will only process files that include the specified table prefix in their names.

## Notes

- This script is designed for a specific BPS data structure. Ensure your Excel files match the expected format.
- The script assumes certain column names and structures. Modify as needed for different data formats.
- Only files containing the specified table prefix in their names will be processed.

## Contributing

Feel free to fork this repository and submit pull requests with any improvements or bug fixes.

## Author

- **Fajrian Aidil Pratama**
- Email: fajrianaidilp@gmail.com

## License

This project is open-source and available under the [MIT License](https://opensource.org/licenses/MIT).
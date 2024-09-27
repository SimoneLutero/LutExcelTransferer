# LUT Excel Transferer

## Overview

This script allows you to copy specific column values from one Excel sheet to another, based on matching identifier columns. It is designed to handle large datasets and provides feedback through progress bars and colored messages.

## Requirements

- Python 3.x
- `openpyxl` library

To install the required packages, run:

```bash
pip install -r requirements.txt
```

## Usage

The script can be executed via the command line with several arguments. Here is the general format:

```bash
python lut_excel_transferer.py -s <source_filename> -d <dest_filename> -si <source_identifiers> -di <dest_identifiers> -sc <source_columns_to_copy> -sp <dest_columns_to_paste> [-so <source_offset>] [-sl <source_limit>] [-o <output_filename>] [-od <output_dir>]
```

## File Preparation

Before running the script, place the Excel files in the same folder as `lut_excel_transferer.py`. If you do not want to specify the file names in the parameters, rename the files as follows:

- **Source file**: `source.xlsx`
- **Destination file**: `dest.xlsx`

⚠️ Note: The files must have a **single row header**. If they have more or fewer header rows, the script will need to be manually modified (for now).

### Notes

- The source and destination files must have a header row.
- Matching columns must be provided in the same order for both source and destination files.
- The script provides real-time progress updates and uses colored messages to indicate success, warnings, and errors.

### Required Parameters
- `-s` or `--source-filename`: The source Excel file (default is source.xlsx).
- `-d` or `--dest-filename`: The destination Excel file (default is dest.xlsx).
- `-si` or `--source-to-match-identifiers`: List of column titles to match in the source sheet.
- `-di` or `--dest-to-match-identifiers`: List of column titles to match in the destination sheet, in the same order as the source identifiers.
- `-sc` or `--source-to-copy-identifiers`: List of column titles from the source sheet to copy.
- `-sp` or `--dest-to-paste-identifiers`: List of column titles from the destination sheet to paste values, in the same order as the source identifiers.

### Optional Parameters
- `-so` or `--source-offset`: The number of rows to skip from the start of the source sheet (default is 0).
- `-sl` or `--source-limit`: The number of rows to process from the source sheet (default is None, meaning all rows).
- `-o` or `--output-filename`: The name of the output file (default is output.xlsx).
- `-od` or `--output-dir`: The directory to save the output file (default is Output).
- `-deb` or `--debug`: Enable debug messages for more detailed logging.

### Example

To transfer values from one Excel file to another where specific columns match:

```bash
python lut_excel_transferer.py -s 'source.xlsx' -d 'dest.xlsx' -si 'Date' 'Number' -di 'Date_Num' 'Number_ID' -sc 'Code' -sp 'Code_ID' -sl 1000
```

In this example:
- The script matches rows where the `Date` and `Number` columns in `source.xlsx` match the `Date_Num` and `Number_ID` columns in `dest.xlsx`.
- It copies the values from the `Code` column in the source sheet to the `Code_ID` column in the destination sheet.
- Only the first 1000 rows of the source sheet are processed.

### Output

The results are saved in the `Output` directory (default) or any directory specified via `--output-dir`.

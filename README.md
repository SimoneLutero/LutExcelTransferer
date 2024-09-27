Ecco il file `README.md` completamente in inglese:

# LUT Excel Transferer

## Python Version

- Python 3.10.0

## Installing required packages

Before running the script, make sure to install the required packages using the following command:

```bash
pip install -r .\requirements.txt
```

## Usage

For information on how to use the script and the available parameters, run:

```bash
python .\lut_excel_transferer.py -h
```

## File Preparation

Before running the script, place the Excel files in the same folder as `lut_excel_transferer.py`. If you do not want to specify the file names in the parameters, rename the files as follows:

- **Source file**: `source.xlsx`
- **Destination file**: `dest.xlsx`

⚠️ Note: The files must have a **single row header**. If they have more or fewer header rows, the script will need to be manually modified (for now).

## Example Usage

```bash
python .\lut_excel_transferer.py -s 'source.xlsx' 'dest.xlsx' -si 'NAME' 'LAST_NAME' -di 'CLIENT_NAME' 'CLIENT_LASTNAME' -sc 'ID' -sp 'CLIENT_ID'
```

### Explanation

This script compares rows from two Excel files:

1. **Source file**: `source.xlsx`
2. **Destination file**: `dest.xlsx`

The script searches for rows where:

- The `NAME` column from the source file matches the `CLIENT_NAME` column from the destination file.
- The `LAST_NAME` column from the source file matches the `CLIENT_LASTNAME` column from the destination file.

Then, it copies the value from the `ID` column of the source file to the `CLIENT_ID` column of the destination file for the matching rows.

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
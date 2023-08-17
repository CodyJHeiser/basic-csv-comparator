import pandas as pd
import time
import re
from os import listdir
from os.path import isfile, join


def cleanse_column_data(data):
    ''' Remove all non-alphanumeric characters except numbers and lowercase/uppercase alphabets '''
    if pd.isna(data):  # Handle NaN values
        return data
    str_data = str(data)  # Convert data to string
    return ''.join(re.findall(r'[a-zA-Z0-9]', str_data))


def compare_csv_files(old_file, new_file, columns, exportName=None, cleanse_data=True):
    if isinstance(columns, str):
        columns = [columns]

    # Read in the csv files
    old_df = pd.read_csv(old_file, sep=',', error_bad_lines=False, dtype='unicode')
    new_df = pd.read_csv(new_file, sep=',', error_bad_lines=False, dtype='unicode')

    # Ensure columns exist in both DataFrames
    for col in columns:
        assert col in old_df.columns, f"'{col}' not found in {old_file}"
        assert col in new_df.columns, f"'{col}' not found in {new_file}"

    # Cleanse data if required
    if cleanse_data:
        old_df[columns] = old_df[columns].applymap(cleanse_column_data)
        new_df[columns] = new_df[columns].applymap(cleanse_column_data)

    # Set indices and match column order33
    old_df.set_index(columns, drop=True, inplace=True)
    new_df.set_index(columns, drop=True, inplace=True)
    old_df = old_df[new_df.columns]

     # Create side-by-side comparison DataFrame
    oldSuffixName=old_file.split('\\').pop().split('.')[0]
    newSuffixName=new_file.split('\\').pop().split('.')[0]
    comparison_df = pd.concat([old_df.add_suffix(f"_{oldSuffixName}"), new_df.add_suffix(f"_{newSuffixName}")], axis=1)

    # Interleave columns for a side-by-side view
    cols = [item for pair in zip(old_df.columns + f"_{oldSuffixName}", new_df.columns + f"_{newSuffixName}") for item in pair]
    comparison_df = comparison_df[cols]

    # Identify mismatched rows for side-by-side comparison
    mismatched_rows = ~old_df.eq(new_df).all(axis=1)
    row_no_match_df = comparison_df.loc[mismatched_rows]

    # Determine export file name
    export_filename = 'export/{}.xlsx'.format(exportName if exportName else str(time.time()).split('.')[0])

    # Excel writer setup
    writer = pd.ExcelWriter(export_filename, engine='xlsxwriter')
    workbook = writer.book
    red_format = workbook.add_format({'bg_color': 'red'})   

    # Write DataFrames to Excel
    old_df[mismatched_rows].reset_index().to_excel(writer, sheet_name='Row Matched', index=False)
    row_no_match_df.reset_index().to_excel(writer, sheet_name='Row Did Not Match', index=False)
    
    # Highlight non-matching cells in 'Row Did Not Match' sheet
    worksheet = writer.sheets['Row Did Not Match']
    for row_idx, (_, row) in enumerate(row_no_match_df.iterrows()):
        for col_idx in range(0, len(row) // 2):
            old_cell = row[f"{old_df.columns[col_idx]}_{oldSuffixName}"]
            new_cell = row[f"{new_df.columns[col_idx]}_{newSuffixName}"]
            if old_cell != new_cell:
                worksheet.write(row_idx + 1, 2*col_idx, str(old_cell), red_format)     # old_cell
                worksheet.write(row_idx + 1, 2*col_idx + 1, str(new_cell), red_format) # new_cell

    try:
        writer.save()
    except Exception:
        try:
            writer._save()
        except Exception as e:
            print(f"Error saving file: {e}")

    print(f"Successfully exported to {export_filename}")

# Usage:
# compare_csv_files('old.csv', 'new.csv', 'column_name_to_compare', exportName='report_name')
baseArchitecture = r"basic-csv-comparator\import"
folderNames = []
driller = []

idColumn = "instrumentIdentifier"

for folder in folderNames:
    path = f"{baseArchitecture}\{folder}"
    files = {}

    for drill in driller:
        drilledPath = f"{path}\{drill}"
        file = [f for f in listdir(drilledPath) if isfile(
            join(drilledPath, f))][0]

        files[drill.split()[1]] = (f"{drilledPath}\{file}")

    # Fetch the comparison data
    oldFile = files[""]
    newFile = files[""]
    exportName = f"{folder}-Compared"

    print(f"Comparing {exportName} export file...")
    compare_csv_files(oldFile, newFile, idColumn, exportName=exportName)

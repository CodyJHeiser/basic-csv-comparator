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
    old_df = pd.read_csv(
        old_file, sep=',', error_bad_lines=False, index_col=False, dtype='unicode')
    new_df = pd.read_csv(
        new_file, sep=',', error_bad_lines=False, index_col=False, dtype='unicode')

    # Ensure that the columns exist in both DataFrames
    for col in columns:
        assert col in old_df.columns, f"'{col}' not found in {old_file}"
        assert col in new_df.columns, f"'{col}' not found in {new_file}"

    # Cleanse the column data if cleanse_data is True
    if cleanse_data:
        old_df[columns] = old_df[columns].applymap(cleanse_column_data)
        new_df[columns] = new_df[columns].applymap(cleanse_column_data)

    # Set index for both dataframes
    old_df.set_index(columns, inplace=True)
    new_df.set_index(columns, inplace=True)
    
    # Determine rows with matching indices
    matched_indices = old_df.index.intersection(new_df.index)
    
    row_match_df = pd.DataFrame()
    row_no_match_df = pd.DataFrame()

    # Check each row for matching content
    for idx in matched_indices:
        old_row = old_df.loc[idx]
        new_row = new_df.loc[idx]
        
        if old_row.equals(new_row):
            row_match_df = row_match_df.append(old_df.loc[[idx]])
        else:
            row_no_match_df = row_no_match_df.append(old_df.loc[[idx]])
    
    # Determine export file name
    if exportName:
        export_filename = 'export/{}.xlsx'.format(exportName)
    else:
        unix = str(time.time()).split('.')[0]
        export_filename = 'export/output-{}.xlsx'.format(unix)

    # Create a Pandas Excel writer using XlsxWriter as the engine
    writer = pd.ExcelWriter(export_filename, engine='xlsxwriter')
    workbook = writer.book
    red_format = workbook.add_format({'bg_color': 'red'})

    # Write each DataFrame to a different worksheet
    row_match_df.reset_index().to_excel(writer, sheet_name='Row Matched', index=False)
    row_no_match_df.reset_index().to_excel(writer, sheet_name='Row Did Not Match', index=False)
    
    # Get the worksheet to add formats
    worksheet = writer.sheets['Row Did Not Match']

    for idx, (old_row, new_row) in enumerate(zip(row_no_match_df.iterrows(), new_df.loc[row_no_match_df.index].iterrows())):
        _, old_values = old_row
        _, new_values = new_row
        
        for col_idx, (old_cell, new_cell) in enumerate(zip(old_values, new_values)):
            if old_cell != new_cell:
                worksheet.write(idx + 1, col_idx, str(old_cell), red_format)

    # Close the Pandas Excel writer and output the Excel file
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

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


def compare_csv_files(old_file, new_file, columns, exportName=None, cleanse_data=False):
    if isinstance(columns, str):
        columns = [columns]

    # Read in the csv files
    old_df = pd.read_csv(
        old_file, sep=',', error_bad_lines=False, dtype='unicode')
    new_df = pd.read_csv(
        new_file, sep=',', error_bad_lines=False, dtype='unicode')

    # Ensure columns exist in both DataFrames
    for col in columns:
        assert col in old_df.columns, f"'{col}' not found in {old_file}"
        assert col in new_df.columns, f"'{col}' not found in {new_file}"

    # Cleanse data if required
    if cleanse_data:
        old_df[columns] = old_df[columns].applymap(cleanse_column_data)
        new_df[columns] = new_df[columns].applymap(cleanse_column_data)

    # Set indices and match column order
    old_df.set_index(columns, drop=True, inplace=True)
    new_df.set_index(columns, drop=True, inplace=True)

    # Rows that matched perfectly
    perfectly_matched_rows = old_df.eq(new_df).all(axis=1)

    old_df = old_df[new_df.columns]

    # Create side-by-side comparison DataFrame
    oldSuffixName = old_file.split('\\').pop().split('.')[0]
    newSuffixName = new_file.split('\\').pop().split('.')[0]

    # Append the suffixes to each DataFrame's columns
    old_df = old_df.add_suffix(f"_{oldSuffixName}")
    new_df = new_df.add_suffix(f"_{newSuffixName}")

    not_in_new = old_df.index.difference(new_df.index)
    not_in_old = new_df.index.difference(old_df.index)

    not_matched_to_old_df = old_df.loc[not_in_new]
    new_rows_found_df = new_df.loc[not_in_old]  # .reset_index()

    # Concatenate the DataFrames and interleave columns
    comparison_df = pd.concat([old_df, new_df], axis=1)
    cols = [item for pair in zip(old_df.columns, new_df.columns)
            for item in pair]
    comparison_df = comparison_df[cols]

    # Identify mismatched rows for side-by-side comparison
    mismatched_to_new_rows = ~old_df.eq(new_df).all(axis=1)
    mismatched_to_new_rows = mismatched_to_new_rows.reindex(
        comparison_df.index, fill_value=False)
    mismatched_to_new_rows = mismatched_to_new_rows.reindex(
        old_df.index, fill_value=False)

    mismatched_to_old_rows = ~new_df.eq(old_df).all(axis=1)
    mismatched_to_old_rows = mismatched_to_old_rows.reindex(
        comparison_df.index, fill_value=False)
    mismatched_to_old_rows = mismatched_to_old_rows.reindex(
        new_df.index, fill_value=False)

    perfectly_matched_df = old_df[perfectly_matched_rows]
    old_partially_matched_df = old_df[mismatched_to_new_rows &
                                      mismatched_to_old_rows & ~perfectly_matched_rows]
    new_partially_matched_df = new_df[mismatched_to_new_rows &
                                      mismatched_to_old_rows & ~perfectly_matched_rows]

    # Concatenate the DataFrames and interleave columns
    partially_matched_df = pd.concat(
        [old_partially_matched_df, new_partially_matched_df], axis=1)
    partial_cols = [item for pair in zip(old_partially_matched_df.columns, new_partially_matched_df.columns)
                    for item in pair]
    partially_matched_df = partially_matched_df[partial_cols]

    # Excel writer setup
    export_filename = 'export/{}.xlsx'.format(
        exportName if exportName else str(time.time()).split('.')[0])
    writer = pd.ExcelWriter(export_filename, engine='xlsxwriter')
    workbook = writer.book

    red_format_old = workbook.add_format(
        {'bg_color': 'red', 'font_color': 'white'})  # for old_df
    red_format_new = workbook.add_format(
        {'bg_color': 'blue', 'font_color': 'white'})  # for new_df

    # Good
    perfectly_matched_df.reset_index().to_excel(
        writer, sheet_name='perfectly_matched_rows', index=False)
    partially_matched_df.reset_index().to_excel(
        writer, sheet_name='partially_matched_df', index=False)
    not_matched_to_old_df.reset_index().to_excel(
        writer, sheet_name='not_matched_to_old_df', index=False)
    new_rows_found_df.reset_index().to_excel(
        writer, sheet_name='new_rows_found_df', index=False)

    # Highlight non-matching cells in 'Row Did Not Match' sheet
    worksheet = writer.sheets['partially_matched_df']

    for row_idx, (index, row) in enumerate(partially_matched_df.iterrows()):
        for col_idx, (old_col, new_col) in enumerate(zip(old_df.columns, new_df.columns)):

            old_cell = row[old_col]
            new_cell = row[new_col]

            if pd.isna(old_cell):
                old_cell = ''
            if pd.isna(new_cell):
                new_cell = ''

            if old_cell != new_cell:
                worksheet.write(row_idx + 1, 2*col_idx+1,
                                str(old_cell), red_format_old)  # old_cell
                worksheet.write(row_idx + 1, 2*col_idx+2,
                                str(new_cell), red_format_new)  # new_cell
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
    compare_csv_files(oldFile, newFile, idColumn, exportName=exportName, cleanse_data=False)

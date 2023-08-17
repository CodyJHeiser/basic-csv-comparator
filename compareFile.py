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

    old_df, new_df = _read_csv_files(old_file, new_file, columns)
    if cleanse_data:
        _cleanse_dataframes(old_df, new_df, columns)
    perfectly_matched_rows, mismatched_to_new_rows, mismatched_to_old_rows = _get_comparison_indices(
        old_df, new_df, columns)

    perfectly_matched_df, partially_matched_df, not_matched_to_old_df, new_rows_found_df = _create_comparative_dfs(
        old_file, new_file, old_df, new_df, perfectly_matched_rows, mismatched_to_new_rows, mismatched_to_old_rows)

    _export_to_excel(old_df, new_df, old_file, new_file, perfectly_matched_df, partially_matched_df,
                     not_matched_to_old_df, new_rows_found_df, exportName)


def _read_csv_files(old_file, new_file, columns):
    old_df = pd.read_csv(old_file, sep=',',
                         error_bad_lines=False, dtype='unicode')
    new_df = pd.read_csv(new_file, sep=',',
                         error_bad_lines=False, dtype='unicode')

    for col in columns:
        assert col in old_df.columns, f"'{col}' not found in {old_file}"
        assert col in new_df.columns, f"'{col}' not found in {new_file}"

    return old_df, new_df


def _cleanse_dataframes(old_df, new_df, columns):
    old_df[columns] = old_df[columns].applymap(cleanse_column_data)
    new_df[columns] = new_df[columns].applymap(cleanse_column_data)


def _get_comparison_indices(old_df, new_df, columns):
    old_df.set_index(columns, drop=True, inplace=True)
    new_df.set_index(columns, drop=True, inplace=True)
    perfectly_matched_rows = old_df.eq(new_df).all(axis=1)

    mismatched_to_new_rows = ~old_df.eq(new_df).all(axis=1)
    mismatched_to_old_rows = ~new_df.eq(old_df).all(axis=1)

    return perfectly_matched_rows, mismatched_to_new_rows, mismatched_to_old_rows


def _create_comparative_dfs(old_file, new_file, old_df, new_df, perfectly_matched_rows, mismatched_to_new_rows, mismatched_to_old_rows):
    old_suffix = old_file.split('\\').pop().split('.')[0].replace(" ", "_")
    new_suffix = new_file.split('\\').pop().split('.')[0].replace(" ", "_")

    old_df = old_df[new_df.columns]
    old_df = old_df.add_suffix(f"_{old_suffix}")
    new_df = new_df.add_suffix(f"_{new_suffix}")

    perfectly_matched_df = old_df[perfectly_matched_rows]

    # Find the differences in rows between old and new
    not_matched_to_old_df = old_df.loc[old_df.index.difference(new_df.index)]
    new_rows_found_df = new_df.loc[new_df.index.difference(old_df.index)]

    # Create a set of indices we don't want in the partially matched dataframe
    exclude_indices = set(perfectly_matched_df.index) | set(
        not_matched_to_old_df.index) | set(new_rows_found_df.index)

    # Now we will remove those indices from the old and new DataFrames to get the partial matches
    old_partial_match = old_df.loc[~old_df.index.isin(exclude_indices)]
    new_partial_match = new_df.loc[~new_df.index.isin(exclude_indices)]

    # Now filter out the rows that don't match between the two datasets to get our partially matched dataframes
    mismatched_to_new_rows = ~old_partial_match.eq(
        new_partial_match).all(axis=1)
    mismatched_to_old_rows = ~new_partial_match.eq(
        old_partial_match).all(axis=1)

    old_partially_matched_df = old_df[mismatched_to_new_rows &
                                      ~perfectly_matched_rows]
    new_partially_matched_df = new_df[mismatched_to_old_rows &
                                      ~perfectly_matched_rows]

    # Concatenate the old and new partially matched dataframes
    partially_matched_df = pd.concat(
        [old_partially_matched_df, new_partially_matched_df], axis=1)

    return perfectly_matched_df, partially_matched_df, not_matched_to_old_df, new_rows_found_df


def _export_to_excel(old_df, new_df, old_file, new_file, perfectly_matched_df, partially_matched_df, not_matched_to_old_df, new_rows_found_df, exportName):
    export_filename = f'export/{exportName if exportName else str(time.time()).split(".")[0]}.xlsx'
    with pd.ExcelWriter(export_filename, engine='xlsxwriter') as writer:
        workbook = writer.book

        red_format_old = workbook.add_format(
            {'bg_color': 'red', 'font_color': 'white'})
        red_format_new = workbook.add_format(
            {'bg_color': 'blue', 'font_color': 'white'})

        # Export dataframes to Excel
        perfectly_matched_df.reset_index().to_excel(
            writer, sheet_name='Matching', index=False)
        partially_matched_df.reset_index().to_excel(
            writer, sheet_name='Non-Matching', index=False)
        not_matched_to_old_df.reset_index().to_excel(
            writer, sheet_name='Old-Not-In-New', index=False)
        new_rows_found_df.reset_index().to_excel(
            writer, sheet_name='New-Not-In-Old', index=False)

        _highlight_mismatches(writer, 'Non-Matching', partially_matched_df, old_file, new_file,
                              old_df.columns, new_df.columns, red_format_old, red_format_new)

    print(f"Successfully exported to {export_filename}")


def _highlight_mismatches(writer, sheet_name, partially_matched_df, old_file, new_file, old_columns, new_columns, red_format_old, red_format_new):
    worksheet = writer.sheets[sheet_name]
    old_suffix = old_file.split('\\').pop().split('.')[0].replace(" ", "_")
    new_suffix = new_file.split('\\').pop().split('.')[0].replace(" ", "_")

    for row_idx, (index, row) in enumerate(partially_matched_df.iterrows()):
        for col_idx, (old_col, new_col) in enumerate(zip(old_columns, new_columns)):
            old_cell, new_cell = str(row[f"{old_col}_{old_suffix}"]), str(
                row[f"{new_col}_{new_suffix}"])

            if old_cell != new_cell:
                worksheet.write(row_idx + 1, 2 * col_idx + 1,
                                old_cell, red_format_old)
                worksheet.write(row_idx + 1, 2 * col_idx + 2,
                                new_cell, red_format_new)


def cleanse_column_data(data):
    # Add your cleansing logic here
    return data
    
# Example call
# compare_csv_files("old_file.csv", "new_file.csv", "column_name")
   
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

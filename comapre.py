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


def insert_original_column(df, column):
    ''' Inserts a copy of the given column right next to it with a prefix "original_" '''
    col_index = df.columns.get_loc(column)
    df.insert(col_index, f"original_{column}", df[column])


def cleanse_and_store_original(df, columns):
    for col in columns:
        insert_original_column(df, col)
        df[col] = df[col].apply(cleanse_column_data)


def get_combined_values_for_columns(df, columns):
    return set(df[columns].apply(lambda x: '|'.join(map(str, x)), axis=1))


def compare_csv_files(old_file, new_file, columns, exportName=None, cleanse_data=False):
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
        cleanse_and_store_original(old_df, columns)
        cleanse_and_store_original(new_df, columns)

    # Find new and missing fields
    old_values = get_combined_values_for_columns(old_df, columns)
    new_values = get_combined_values_for_columns(new_df, columns)

    new_fields = new_values - old_values
    missing_fields = old_values - new_values
    matched_fields = old_values.intersection(new_values)  # Capture matched fields

    # Filter based on combined values
    old_df['combined'] = old_df[columns].apply(lambda x: '|'.join(map(str, x)), axis=1)
    new_df['combined'] = new_df[columns].apply(lambda x: '|'.join(map(str, x)), axis=1)

    new_df_filtered = new_df[new_df['combined'].isin(new_fields)]
    old_df_filtered = old_df[old_df['combined'].isin(missing_fields)]
    matched_df = new_df[new_df['combined'].isin(matched_fields)]  # Capture matched rows

    # Determine export file name
    if exportName:
        export_filename = 'export/{}.xlsx'.format(exportName)
    else:
        unix = str(time.time()).split('.')[0]
        export_filename = 'export/output-{}.xlsx'.format(unix)

    # Create a Pandas Excel writer using XlsxWriter as the engine
    writer = pd.ExcelWriter(export_filename, engine='xlsxwriter')

    # Write each DataFrame to a different worksheet
    new_df_filtered.to_excel(writer, sheet_name='New Fields', index=False)
    old_df_filtered.to_excel(writer, sheet_name='Missing Fields', index=False)
    matched_df.to_excel(writer, sheet_name='Matched', index=False)  # Export matched rows

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

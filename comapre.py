import pandas as pd
import time


def compare_csv_files(old_file, new_file, column):
    # Read in the csv files
    old_df = pd.read_csv(old_file)
    new_df = pd.read_csv(new_file)

    # Ensure that the column exists in both DataFrames
    assert column in old_df.columns, f"'{column}' not found in {old_file}"
    assert column in new_df.columns, f"'{column}' not found in {new_file}"

    # Find new and missing fields
    old_values = set(old_df[column])
    new_values = set(new_df[column])
    new_fields = new_values - old_values
    missing_fields = old_values - new_values

    # Create DataFrames for new and missing fields
    new_df_filtered = new_df[new_df[column].isin(new_fields)]
    old_df_filtered = old_df[old_df[column].isin(missing_fields)]

    # Create a Pandas Excel writer using XlsxWriter as the engine
    unix = str(time.time()).split('.')[0]
    writer = pd.ExcelWriter('export/output-{}.xlsx'.format(unix), engine='xlsxwriter')

    # Write each DataFrame to a different worksheet
    new_df_filtered.to_excel(writer, sheet_name='New Fields', index=False)
    old_df_filtered.to_excel(writer, sheet_name='Missing Fields', index=False)

    # Close the Pandas Excel writer and output the Excel file
    writer._save()


# Usage: compare_csv_files('old.csv', 'new.csv', 'column_name_to_compare')
compare_csv_files('import/2006.csv', 'import/2007.csv', 'ga_city')

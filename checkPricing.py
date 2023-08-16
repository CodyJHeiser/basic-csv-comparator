import pandas as pd
import time

def cleanse_cash_price(price):
    ''' Remove all non-numeric characters except for '.' '''
    if pd.isna(price):  # Handle NaN values
        return price
    return str(price).replace(',', '').replace('$', '')

def compare_files(old_file, new_file, exportName=None):
    # Read in the csv files
    old_df = pd.read_csv(old_file)
    new_df = pd.read_csv(new_file)

    # Cleanse the "Cash Price" columns
    old_df['Cash Price'] = old_df['Cash Price'].apply(cleanse_cash_price).astype(float)
    new_df['Cash Price'] = new_df['Cash Price'].apply(cleanse_cash_price).astype(float)

    # Merge dataframes on 'Model' to identify matching models and different prices
    merged = old_df.merge(new_df, on='Model', how='outer', indicator=True, suffixes=('_original', '_website'))

    # Adjust NaN values for Cash Price columns for easy arithmetic operations
    merged['Cash Price_original'].fillna(0, inplace=True)
    merged['Cash Price_website'].fillna(0, inplace=True)

    # Extract rows with matching model but different price (difference is greater than 10)
    mismatches = merged[(merged['_merge'] == 'both') & (abs(merged['Cash Price_original'] - merged['Cash Price_website']) > 10.0)]
    # Add rows where model is not found in the new file
    not_found = merged[merged['_merge'] == 'left_only']
    not_found['Cash Price_website'] = 'N/A'
    mismatches = pd.concat([mismatches, not_found])

    # Extract exact matches (both model and price within the 10 units range)
    exact_matches = merged[(merged['_merge'] == 'both') & (abs(merged['Cash Price_original'] - merged['Cash Price_website']) <= 10.0)][['Model', 'Cash Price_original']]
    
    # Rename columns for exact_matches
    exact_matches.columns = ['Model', 'Cash Price']

    # Determine export file name
    if exportName:
        export_filename = f'export/{exportName}.xlsx'
    else:
        unix = str(time.time()).split('.')[0]
        export_filename = f'export/output-{unix}.xlsx'

    # Create a Pandas Excel writer using XlsxWriter as the engine
    writer = pd.ExcelWriter(export_filename, engine='xlsxwriter')

    # Write each DataFrame to a different worksheet
    exact_matches.to_excel(writer, sheet_name='Matching', index=False)
    mismatches[['Model', 'Cash Price_original', 'Cash Price_website']].to_excel(writer, sheet_name='Not Matching', index=False)

    # Close the Pandas Excel writer and output the Excel file
    writer._save()

# Usage:
compare_files('import/original.csv', 'import/web.csv', 'Price_Missing')

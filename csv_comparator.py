import pandas as pd
import numpy as np
import re  # Regular expression library

def clean_text(text):
    # Convert to string, trim, and remove non-alphanumeric characters
    fixed = re.sub(r'[^A-Za-z0-9]', '', str(text).strip())

    # Convert empty strings to 0
    if fixed == '':
        fixed = '0'

    return fixed

def compare_csv(file1, file2, id_column):
    # Read the CSV files
    df1 = pd.read_csv(file1)
    df2 = pd.read_csv(file2)

    # Renaming columns for clarity
    df1 = df1.add_suffix('_file1')
    df2 = df2.add_suffix('_file2')
    
    # Resetting the ID columns to original name for merging
    df1 = df1.rename(columns={id_column + '_file1': id_column})
    df2 = df2.rename(columns={id_column + '_file2': id_column})

    # Merging the dataframes on the ID column with an indicator
    merged_df = pd.merge(df1, df2, on=id_column, how='outer', indicator=True)

    # Interlace columns
    columns = [id_column]  # Initialize with the ID column
    for col in df1.columns:
        if col != id_column:
            file2_col = col.replace('_file1', '_file2')
            columns.append(col)  # Add file1 column
            if file2_col in df2.columns:
                columns.append(file2_col)  # Add file2 column

    merged_df = merged_df[columns + ['_merge']]  # Include the merge indicator for row status

    # Function to determine the row status
    def row_status(row):
        if row['_merge'] == 'both':
            for i in range(1, len(columns), 2):  # Skip the ID column
                file1_col = columns[i]
                file2_col = columns[i + 1] if i + 1 < len(columns) else None
                file1_val = clean_text(row[file1_col])
                file2_val = clean_text(row[file2_col]) if file2_col else None
                if file1_val != file2_val:
                    return 'not matched'
            return 'matched'
        else:
            return 'missing'

    # Add a status column
    merged_df['Status'] = merged_df.apply(row_status, axis=1)
    # Drop the '_merge' column
    merged_df.drop(columns=['_merge'], inplace=True)

    # Move the 'Status' column to the first position
    cols = ['Status'] + columns
    merged_df = merged_df[cols]

    # Function to apply the highlighting
    def highlight_diff(row):
        colors = {col: '' for col in cols}  # Initialize with empty strings
        for i in range(2, len(cols), 2):  # Start from 2 to skip the 'Status' and ID columns
            file1_col = cols[i]
            file2_col = cols[i + 1] if i + 1 < len(cols) else None
            if file2_col:
                file1_val = clean_text(row[file1_col])
                file2_val = clean_text(row[file2_col])
                color = 'background-color: yellow' if file1_val != file2_val else ''
                colors[file1_col] = color
                colors[file2_col] = color
        return pd.Series(colors)

    # Applying the highlight function
    styled_df = merged_df.style.apply(highlight_diff, axis=1)

    # Saving the result to an Excel file (as CSV doesn't support styling)
    styled_df.to_excel('output.xlsx', engine='openpyxl', index=False)


# Example usage
file1 = "import/test.csv"
file2 = "import/test_diff.csv"

compare_csv(file1, file2, 'ID')
print("Done!")
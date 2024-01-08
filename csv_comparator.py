import pandas as pd
import numpy as np
import re  # Regular expression library

def clean_text(text):
    # Convert to string, trim, and remove non-alphanumeric characters
    return re.sub(r'[^A-Za-z0-9]', '', str(text).strip())

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

    # Merging the dataframes on the ID column
    merged_df = pd.merge(df1, df2, on=id_column, how='outer')

    # Function to apply the highlighting
    def highlight_diff(row):
        colors = {}
        for col in df1.columns:
            if col != id_column:
                file1_val = clean_text(row[col])
                file2_col = col.replace('_file1', '_file2')
                file2_val = clean_text(row[file2_col]) if file2_col in merged_df.columns else None
                color = 'background-color: yellow' if file1_val != file2_val else ''
                colors[col] = color
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
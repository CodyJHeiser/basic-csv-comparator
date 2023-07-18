# CSV File Comparison Tool

This tool is designed to compare a specific column in two CSV files and output the differences in two separate Excel sheets.

## Dependencies

This script uses the following Python libraries:

- pandas
- xlsxwriter

You can install these using pip:

```bash
pip install pandas xlsxwriter
```

## Usage

The script requires three arguments:

1. The name of the old CSV file
2. The name of the new CSV file
3. The name of the column to compare

Here is an example of how to use the script:

```python
compare_csv_files('old.csv', 'new.csv', 'column_name_to_compare')
```

In this example, 'old.csv' and 'new.csv' are the old and new CSV files respectively, and 'column_name_to_compare' is the name of the column to be compared.

## Output

The script will output an Excel file named 'output.xlsx', which contains two sheets:

- 'New Fields': contains the rows from the new CSV file where the specified column contains values that do not exist in the old CSV file.
- 'Missing Fields': contains the rows from the old CSV file where the specified column contains values that do not exist in the new CSV file.

```

Please replace `'old.csv'`, `'new.csv'`, and `'column_name_to_compare'` with your actual old file name, new file name, and the column name to be compared, respectively.
```

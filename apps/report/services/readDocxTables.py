import pandas as pd
from docx import Document

# Replace 'your_file.docx' with the path to your .docx file
document = Document('files/WAIS.docx')

# Retrieve all tables in the document
tables = document.tables

# List to hold all DataFrames
dataframes = []

for idx, table in enumerate(tables):
    data = []
    keys = None
    for i, row in enumerate(table.rows):
        # Extract text from each cell in the row
        text = [cell.text.strip() for cell in row.cells]
        if i == 0:
            # Assume the first row contains column headings
            keys = text
        else:
            data.append(text)
    # Create DataFrame for the current table
    df = pd.DataFrame(data, columns=keys)
    dataframes.append(df)

    # Optionally, save each DataFrame to a CSV file
    df.to_csv(f"table_{idx+1}.csv", index=False)

# Now, 'dataframes' is a list containing a DataFrame for each table
# You can access them individually, for example:
# dataframes[0] for the first table, dataframes[1] for the second, etc.

# Example: Print all DataFrames
for i, df in enumerate(dataframes):
    print(f"Table {i+1}:")
    print(df)
    print("\n")

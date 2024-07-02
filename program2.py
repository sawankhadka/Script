import pandas as pd

# Read the dataset
df = pd.read_excel("22BCAA54_OUTPUT1.xlsx")

# Ensure the 'Obt Marks' column is of type object (string) to avoid dtype issues
df['Obt Marks'] = df['Obt Marks'].astype(object)
df.loc[(df['Exam Type'].isin(['Aggregate', 'SA'])) & (df['Obt Marks'] == 0), 'Obt Marks'] = ''

# Pivot the table to get the desired format
pivot_df = df.pivot_table(index=['Roll No', 'Student Name', 'Batch Name', 'Subject Name'], 
                          columns='Exam Type', values='Obt Marks', aggfunc='sum').reset_index()

# Reorder the columns to have 'FA' and 'SA' first, followed by 'Aggregate'
columns_order = ['Roll No', 'Student Name', 'Batch Name', 'Subject Name', 'FA', 'SA', 'Aggregate']
pivot_df = pivot_df.reindex(columns=columns_order)

# Create an Excel writer object
output_filename = "OrganizedData_with_Filters1.xlsx"
writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')

# Write the pivot table to the Excel file
pivot_df.to_excel(writer, sheet_name='Sheet1', index=False, startrow=1)

# Get the xlsxwriter workbook and worksheet objects.
workbook = writer.book
worksheet = writer.sheets['Sheet1']

# Define a format for the merged cells
merge_format = workbook.add_format({'align': 'center', 'bold': True, 'border': 1})

# Merge cells for the 'Exam Type' column header
worksheet.merge_range('E1:G1', 'Exam Type', merge_format)

# Write column headers
for col_num, value in enumerate(pivot_df.columns):
    worksheet.write(1, col_num, value, merge_format)

# Add an autofilter to the worksheet
worksheet.autofilter(1, 0, pivot_df.shape[0] + 1, pivot_df.shape[1] - 1)

# Add a dropdown filter for the Exam Type columns
for col_num in range(4, 7):  # Columns E (FA), F (SA), G (Aggregate)
    worksheet.data_validation(2, col_num, pivot_df.shape[0] + 1, col_num, {
        'validate': 'list',
        'source': ['FA', 'SA', 'Aggregate']
    })

# Close the Excel writer and save the file
writer.close()

print(f"Data written to '{output_filename}' with filters and 'Exam Type' column grouping.")
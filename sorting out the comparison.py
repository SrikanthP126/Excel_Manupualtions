import pandas as pd

# Load the disposition data (assuming only one sheet per file for simplicity)
disposition_file = 'path_to_your_disposition_file.xlsx'  # Replace with your file path
disposition_df = pd.read_excel(disposition_file, sheet_name=None)

# Load the deleted objects data (assuming only one sheet per file for simplicity)
deleted_objects_file = 'path_to_your_deleted_objects_file.xlsx'  # Replace with your file path
deleted_objects_df = pd.read_excel(deleted_objects_file, sheet_name=None)

# For each sheet in the disposition file, find and remove deleted objects
for sheet_name in disposition_df.keys():
    disposition_sheet = disposition_df[sheet_name]
    
    # Assuming 'Object Name' column exists in both files and is the key
    # Perform an anti-join: keep rows in disposition_sheet that are not in deleted_objects_sheet
    cleaned_disposition_sheet = disposition_sheet[~disposition_sheet['Object Name'].isin(deleted_objects_df[sheet_name]['Object Name'])]
    
    # Save the cleaned sheet back to an Excel file (new file or overwrite)
    with pd.ExcelWriter('cleaned_disposition_file.xlsx', mode='a', engine='openpyxl') as writer:
        cleaned_disposition_sheet.to_excel(writer, sheet_name=sheet_name, index=False)

print("Process completed. Check the 'cleaned_disposition_file.xlsx' for the results.")

import os
import pandas as pd

def list_folders(directory):
    folder_paths = []
    folder_names = []
    for root, dirs, files in os.walk(directory):
        for dir in dirs:
            folder_path = os.path.join(root, dir)
            folder_path = folder_path.replace('/', '\\')  # Convert to backslashes
            folder_paths.append(folder_path)
            folder_names.append(dir)
    return folder_paths, folder_names

def list_folders_to_excel(parent_folder, output_excel_file):
    # Get the list of folder paths and names
    folder_paths, folder_names = list_folders(parent_folder)

    # Create a DataFrame
    df = pd.DataFrame({
        "Folder Path": folder_paths,
        "Folder Name": folder_names
    })

    # Write the DataFrame to an Excel file
    df.to_excel(output_excel_file, index=False)

# Usage example
parent_folder = "//Server/Docs/SW_Requests".replace('/', '\\')
output_excel_file = "C:/Users/Ryan_Lazar/Documents/Scripting/folders_list.xlsx".replace('/', '\\')
list_folders_to_excel(parent_folder, output_excel_file)



import pandas as pd
from pathlib import Path
import shutil

# Load the Excel file
path = "./finalprojectexcel.xlsx"
xl = pd.read_excel(path)

# Ask for a main folder name
folname = input("Please enter the name of the main folder: ")
main_folder = Path(folname)
main_folder.mkdir(exist_ok=True)
print(f"Folder '{main_folder}' created successfully!")

# Get all unique file extensions
extensions = xl['extension'].unique()

# Create folders for each file type & move files
for ext in extensions:
    ext_folder = main_folder / ext  # Create folder path inside main folder
    ext_folder.mkdir(exist_ok=True)  # Make sure it exists
    
    # Filter files with this extension
    filtered_data = xl[xl['extension'] == ext]
    
    for _, row in filtered_data.iterrows():
        file_name = row['file name']  # Assuming file names are in this column
        file_path = Path(file_name)
        
        if not file_path.exists():  # If file doesn't exist, create it
            file_path.touch()
            print(f"Created empty file: {file_name}")

        # Move file to the correct folder
        shutil.move(str(file_path), str(ext_folder / file_path.name))
        print(f"Moved '{file_name}' to '{ext_folder}'")

print("All files have been created and organized successfully!")
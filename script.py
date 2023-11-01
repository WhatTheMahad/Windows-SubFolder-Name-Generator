import os
import openpyxl

# Function to get the names of all folders in a directory
def get_folder_names(directory):
    folder_names = []
    for root, dirs, files in os.walk(directory):
        for folder in dirs:
            folder_names.append(folder)
    return folder_names

# Function to create an Excel file and write folder names to it
def create_excel_with_folder_names(directory, output_excel_file):
    folder_names = get_folder_names(directory)

    # Create a new Excel workbook
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Folder Names"

    # Write folder names to the Excel sheet
    for i, folder_name in enumerate(folder_names):
        sheet.cell(row=i + 1, column=1, value=folder_name)

    # Save the Excel file
    workbook.save(output_excel_file)
    print(f"Excel file '{output_excel_file}' created successfully.")

if __name__ == "__main__":
    folder_path = r'G:\Seasons n Movies\Movies\All of it'  # Replace with the path to your folder
    excel_file = 'folder_names.xlsx'  # Output Excel file name

    create_excel_with_folder_names(folder_path, excel_file)

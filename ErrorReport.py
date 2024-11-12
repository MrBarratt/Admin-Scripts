import os
import pandas as pd

# Define paths
source_folder = r'C:\Users\USER\OneDrive - XpertEase\Imports-Errors\Error'
completed_folder = os.path.join(source_folder, 'Completed')

# Create Completed folder if it doesn't exist
if not os.path.exists(completed_folder):
    os.makedirs(completed_folder)

# Prepare to store processed file names
processed_files = []

# Process each txt file in the source folder
for file_name in os.listdir(source_folder):
    if file_name.endswith('.txt'):
        file_path = os.path.join(source_folder, file_name)
        
        # Read the text file into a list of lines
        with open(file_path, 'r', encoding='utf-8') as file:
            lines = file.readlines()
        
        # Filter out the unwanted rows (case insensitive)
        filtered_lines = [
            line.strip() for line in lines 
            if not line.startswith('File:') and 'client program not found' not in line.lower()
        ]
        
        # Convert the filtered lines to a pandas DataFrame with one column
        data = pd.DataFrame(filtered_lines, columns=['Data'])
        
        # Create an Excel file with the same name as the text file
        excel_filename = os.path.splitext(file_name)[0] + '.xlsx'
        excel_filepath = os.path.join(completed_folder, excel_filename)

        # Save the DataFrame to an Excel file
        with pd.ExcelWriter(excel_filepath) as writer:
            # Move rows that start with "Phone" to the "Phone" sheet
            phone_data = data[data['Data'].str.startswith('Phone')]
            other_data = data[~data['Data'].str.startswith('Phone')]

            # Write the "Phone" data to the first sheet
            phone_data.to_excel(writer, index=False, header=False, sheet_name='Phone')

            # Write the "Other" data to the second sheet
            other_data.to_excel(writer, index=False, header=False, sheet_name='Other')

        # Add the file name to the processed list
        processed_files.append(file_name)
        
        print(f"Processed and saved: {excel_filepath}")

        # Delete the txt file after processing
        os.remove(file_path)
        print(f"Deleted: {file_path}")

# Create a Completed.txt file in the source folder and write the list of processed file names
completed_txt_path = os.path.join(source_folder, 'Completed.txt')
with open(completed_txt_path, 'w', encoding='utf-8') as completed_file:
    for file_name in processed_files:
        completed_file.write(file_name + '\n')

print(f"Completed file names saved to: {completed_txt_path}")
print("Processing completed and txt files deleted.")

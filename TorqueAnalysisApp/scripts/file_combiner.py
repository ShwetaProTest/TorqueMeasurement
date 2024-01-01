#!/usr/bin/env python
# coding: utf-8

########################################################################################################
# File combiner
# created by:      RHM
# purpose:         Combines input and result files into "Ready_Files" with specific sheet names.
# Created on:      02-11-2023
# version:         1.0
########################################################################################################

import os
import pandas as pd
from config import input_path, result_path, ready_path

class FileProcessor:
    def __init__(self):
        self.input_folder = input_path
        self.results_folder = result_path
        self.ready_folder = ready_path

    def create_ready_folders(self):
        """
        Create the 'Ready_Files' folder if it doesn't exist.
        """
        if not os.path.exists(self.ready_folder):
            os.makedirs(self.ready_folder)

    def process_files(self):
        """
        Process input and result files, combining them into 'Ready_Files' with sheet names 'raw_data' and 'raw_analysed_data'.
        """
        # Get a list of input files in the 'Input' folder
        input_files = [file for file in os.listdir(self.input_folder) if file.lower().endswith(".xlsx")]

        # Iterate through the input files
        for input_file in input_files:
            input_filename, input_extension = os.path.splitext(input_file)

            # Check if the same file name exists in the 'Result' folder
            result_folder_path = os.path.join(self.results_folder, input_filename)
            if os.path.exists(result_folder_path):
                result_file_name = f"Result_{input_filename}.xlsx"
                result_file_path = os.path.join(result_folder_path, result_file_name)

                if os.path.exists(result_file_path):
                    print(f"Matching result file for '{input_file}' found in the 'Result' folder.")

                    # Read the input and result Excel files into dataframes
                    input_df = pd.read_excel(os.path.join(self.input_folder, input_file))
                    result_df = pd.read_excel(result_file_path)

                    # Create the folder structure in 'Ready_Files'
                    ready_file_folder = os.path.join(self.ready_folder, input_filename)
                    if not os.path.exists(ready_file_folder):
                        os.makedirs(ready_file_folder)

                    # Create a new Excel file in the respective folder within 'Ready_Files'
                    ready_file_name = f"Ready_{input_filename}.xlsx"
                    ready_file_path = os.path.join(ready_file_folder, ready_file_name)

                    # Combine the data into a new Excel file with 'raw_data' and 'raw_analysed_data'
                    with pd.ExcelWriter(ready_file_path, engine='xlsxwriter') as writer:
                        input_df.to_excel(writer, sheet_name='raw_data', index=False)
                        result_df.to_excel(writer, sheet_name='raw_analysed_data', index=False)

                    print(f"Combined files for '{input_file}' saved as '{ready_file_name}' in the 'Ready_Files/{input_filename}' folder.")
                else:
                    print(f"No matching result file for '{input_file}' found in the 'Result/{input_filename}' folder.")
            else:
                print(f"The folder '{input_filename}' does not exist in the 'Result' folder.")

if __name__ == "__main__":
    # Initialize the file processor
    processor = FileProcessor()
    
    # Create the 'Ready_Files' folder if it doesn't exist
    processor.create_ready_folders()
    
    # Process input and result files and combine them into the 'Ready_Files' folder
    processor.process_files()
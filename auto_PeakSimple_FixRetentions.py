# This script was created to automate the task of creating .CPT for each run
# Instruments: GC-FID
# Software: works with Peak Simple and MS Excel
# Authors: Davi de Ferreyro Monticelli, iREACH group (University of British Columbia)
# Date: 2024-04-22
# Version: 1.0.0

# How it works: The following test tries to SHIFT all the terpenes resolution based on a REF file

#######################################################################################################################
# NOTE: This script was originally designed for organizing the data of 22 terpene species (from calibration)
#       If you are working with different species, please modify where appropriate.

# NOTE: Use the file path and name of your machine/preference.

# NOTE: This script works for LOG files in which the naming convention of original .CHR files is:
#       "Terpenes_Sampling_Zmin_XX_XX_XX_CCC-CCC_STARTXXXX_SXX", where X are numbers,
#       XX_XX_XX is the sampling date, STARTXXXX the start time, and SXX, the sample #,
#       C are characters. If the naming convention differs, please modify the script
#       accordingly. Z is the trapping time in minutes.
#######################################################################################################################


import pandas as pd
import os

# Get the current user's username
username = os.getlogin()

# Specify REFERENCE file path
# This is the .CPT reference file where you replace the ".CPT" by ".txt" using a software such as Notepad++
file_path = f'C:\\Users\\{username}\\PycharmProjects\\GC-FID\\'  # Example
Trap_1min_REF_file_name = 'Trap_1min_Calibration_REF_23_08_08_S01.txt'  # Example

# Read the text file into a DataFrame
# Replace ',' with your actual delimiter if it's different
ref_1min = pd.read_csv(file_path+Trap_1min_REF_file_name, delimiter=',', header=None)

# Specify file paths (the ones to be modified)
file_path = f'C:\\Users\\{username}\\PycharmProjects\\GC-FID\\'  # Example
input_path = os.path.join(file_path, 'input\\')
output_path = os.path.join(file_path, 'output\\')

# Read the CSV file into a DataFrame
file_name = 'PeakSimple_Log.csv'
df = pd.read_csv(file_path+input_path+file_name, delimiter=',')
# Remove columns using drop()
df.drop(df.columns[5:], axis=1, inplace=True)
# Renaming columns
df.columns = ['File', 'Sampling date' 'Sampling hour' 'Terpene' 'Retention time']
# Get the reference file retention time:
# Trap_1min_Calibration_REF_23_08_08_S01 -> Terpenes_Sampling_1min_23_08_08_S01.CHR
file_to_find = "Terpenes_Sampling_1min_23_08_08_S01.CHR"
value_to_subtract = df.loc[df['File'] == file_to_find, 'Retention time'].values[0]
df['Shift'] = df['Retention time'] - value_to_subtract

# The final dataframe should be something like this: - essentially is the LOG file for just a-pinene w/ an added column
# 'File' 'Sampling date' 'Sampling hour' 'Terpene' 'Retention time' Shift
# Terpenes_Sampling_1min_23_07_17_S01.CHR 2023-07-17 13:36:29 a-pinene 8.183 0.02
# Terpenes_Sampling_1min_23_07_17_S02.CHR 2023-07-17 15:05:55 a-pinene 8.163 0
# Terpenes_Sampling_1min_23_07_17_S03.CHR 2023-07-17 16:37:55 a-pinene 8.13 -0.033
print(df)

a_pinene_retention = df

# Iterate through each row of the DataFrame
for index, row in a_pinene_retention.iterrows():
    filename = row['File']
    shift_value = row['Shift']

    # Replace .CHR by .CPT in filename
    filename = filename.replace(".CHR", ".CPT")
    print("Processing resolution for file ", filename)

    # Determine which reference DataFrame to use based on filename
    ref_df = ref_1min.copy()

    # Multiply shift value by 1000
    shift_value = shift_value*1000

    # Add shift value to columns 2 and 3
    ref_df.iloc[:, 2] += shift_value
    ref_df.iloc[:, 3] += shift_value

    # Save the modified DataFrame as a new text file
    output_filename = os.path.join(output_path, filename)
    # Write DataFrame to CSV file with additional control characters
    with open(output_filename, 'wb') as file:
        for _, line in ref_df.iterrows():
            csv_row = ','.join(map(str, line)).encode('utf-8') + b'\r'  # Add CR control character at the end of row
            file.write(csv_row)  # Write the row to the file
            file.write(b'\r\n')  # Add CR LF control character between rows

    # Open the saved text file in binary mode, insert "<TYPE>=COMPONENT" in the first line
    with open(output_filename, 'rb+') as file:
        contents = file.read()
        file.seek(0, 0)
        file.write("<TYPE>=COMPONENT\r\n".encode('utf-8') + contents)  # Add CR LF control character

import pandas as pd

# Define the folder path
folder_path = 'C:\\Users\\Shweta\\PycharmProjects\\exceltrans\\'

# Read the first Excel file from the specified folder into DataFrame
df1 = pd.read_excel(folder_path + 'Input_file.XLSX')

# Read the second Excel file which doesn't have a header
df2 = pd.read_excel(folder_path + 'Result_file.XLSX', header=None)
df2.columns = ['Time']

# Initialize a Temperature column in df2
df2['Temperature'] = None

# Sort df2 based on Time
df2 = df2.sort_values(by='Time')

# For each integer time value in df1
for _, row in df1.iterrows():
    if pd.notna(row['Time']):
        time_int = int(row['Time'])
        # Get the earliest float time value in df2 that starts with the integer
        match_row = df2[df2['Time'].astype(int) == time_int].head(1)

        if not match_row.empty:
            match_index = match_row.index[0]
            df2.at[match_index, 'Temperature'] = row['Temperature']
    else:
        print(f"Skipped row due to NaN value: {row}")

# Fill missing values in the Temperature column with the last valid value
df2['Temperature'] = df2['Temperature'].ffill()

# Write the modified df2 DataFrame to a new CSV file in the same folder
df2.to_csv(folder_path + 'Ready_frequency_sweep_input.csv', index=False)

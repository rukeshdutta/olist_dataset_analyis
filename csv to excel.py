import glob
import pandas as pd

# Define the directory where the CSV files are located
csv_dir = "D:\\Python Projects\\e-commerce data kaggle\\olist_dataset\\"

# Define the output XLSX file name
xlsx_file = "D:\\Python Projects\\e-commerce data kaggle\\olist_dataset\\output.xlsx"

# Create an empty list to store the DataFrames and sheet names
dfs = []
sheet_names = []

# Create an empty list to store the DataFrames and sheet names
dfs = []
sheet_names = []

# Loop through each CSV file in the directory and add it to the list of DataFrames
for csv_file in glob.glob(csv_dir + '*.csv'):
    # Extract the sheet name from the file name
    sheet_name = csv_file.split('\\')[-1].split('.')[0]
    print(sheet_name)
    # Read in the CSV file as a Pandas DataFrame
    df = pd.read_csv(csv_file)
    
    # Append the DataFrame and sheet name to their respective lists
    dfs.append(df)
    sheet_names.append(sheet_name)

# Create a Pandas Excel writer object
writer = pd.ExcelWriter(xlsx_file, engine='xlsxwriter')

# Loop through each DataFrame and sheet name in the lists and write them to the XLSX file
for df, sheet_name in zip(dfs, sheet_names):
    df.to_excel(writer, sheet_name=sheet_name, index=False)

# Save the XLSX file
writer.save()

print('CSV files saved to {} as multiple tabs.'.format(xlsx_file))

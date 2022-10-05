# importing the required modules
import glob
import pandas as pd

# specify the folder path
path = r"C:/path/to/all/files"

# xls files in the path - can be changed to csv or xlsx if needed
file_list = glob.glob(path + "/*.xls")

# list of excel files we want to merge.
# pd.read_excel(file_path) reads the
# excel data into pandas dataframe.
excl_list = []

for file in file_list:
    # Option 1 - Uncomment to use
    # Select sheet number in the workbook to use
    # df = pd.read_excel(file, sheet_name=3)
    
    # Option 2 - Uncomment to use
    # Reads the first sheet in the workbook and uses it
    # df = pd.read_excel(file)
    
    # Option 3 - Uncomment to use
    # Reads the specific sheet name in the workbook
    # df = pd.read_excel(file, 'Sheet1')

    # Option 4 - Uncomment to use
    # Specify the name of the sheet to use in the workbook alt
    # df = pd.read_excel(file, sheet_name='Sheet2')

    df = pd.read_excel(file, 'SheetName')
    df.insert(0, 'new_colum', str(file.split("\\")[len(file.split("\\"))-1]))
    excl_list.append(df)

# concatenate all DataFrames in the list
# into a single DataFrame, returns new
# DataFrame.
excl_merged = pd.concat(excl_list, ignore_index=True)

# exports the dataframe into excel file
# with specified name.
excl_merged.to_excel('Combined_NEW.xlsx', index=False)

import win32com.client


parser = argparse.ArgumentParser(description="inputs for script")
parser.add_argument("--directory", help="Provide the path to the source json file.", required=True,)
    
args = parser.parse_args()
directory_location = args.directory

# Create an instance of Excel
excel = win32com.client.Dispatch('Excel.Application')

# Open the Excel file
workbook = excel.Workbooks.Open(directory_location)

# Do something with the file (e.g. read or modify data)
#...
changes = workbood.active
changes["A1"] = "Hello, World!"
changes["B2"] = 42

# Save the file
workbook.Save()

# Close the file
workbook.Close()

# Release the Excel object
excel.Quit()

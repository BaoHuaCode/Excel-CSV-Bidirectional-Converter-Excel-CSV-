from pathlib import Path
import openpyxl, csv

path = Path(__file__).parent
output_folder = path/'excel'
output_folder.mkdir(exist_ok=True)

# Get all csv files in the folder
csv_files = [file for file in path.iterdir() if file.is_file() and file.suffix=='.csv']

# Create a dict and set the same workbook with it sheets in a form
file_groups = {}
# Loop the files and use file stem split to set excel name and sheet name
for csv_file in csv_files:
    parts = csv_file.stem.split('_', 1)
    if len(parts) == 2:
        excel_name, sheet_name = parts
    else:
        excel_name, sheet_name = csv_file.stem, 'Sheet1'
    
    # Set the dict,excel name as key , in the same workbook sheet name and its csv file 
    # path as tuple keys append in the list key
    file_groups.setdefault(excel_name,[]).append((sheet_name, csv_file))

# Set the excel workbook
for excel_name, sheets in file_groups.items():
    # Create a workbook
    wb = openpyxl.Workbook()
    default_sheet = wb.active
    # Set a flag
    is_first_sheet = True
    print(f"Creating Excel {excel_name}.xlsx...")

    for sheet_name,csv_path in sheets:
        # If is first sheet set the default sheet as it ,and set the title
        if is_first_sheet:
            sheet = default_sheet
            sheet.title = sheet_name
            is_first_sheet = False
        else:
            # if not the first sheet create a new sheet
            wb.create_sheet(title=sheet_name)
        
        # Open a csv file and read it
        with open(csv_file, 'r', encoding='utf-8') as f:
            reader = csv.reader(f)
            # Loop the reader get each row
            for row in reader:
                # set int and float back 
                cleaned_row = []
                for item in row:
                    try:
                        if '.' in item:
                            cleaned_row.append(float(item))
                        else:
                            cleaned_row.append(int(item))
                    # If not number will raise error ,use try except to add str to list
                    except ValueError:
                        cleaned_row.append(item)
                
                # write the rows in excel sheet
                sheet.append(cleaned_row)
    # save the excel
    wb.save(output_folder/f"{excel_name}_recovered.xlsx")
print("All csv files merged and recovered successfully.")


        

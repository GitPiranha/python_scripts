import os
import openpyxl

# folder location
folder_location = "C:\\Test"

# search term
search_term = "caro"

# list of files in folder
list_of_files_in_folder = os.listdir(folder_location)
print(list_of_files_in_folder)

list_of_full_paths = []

for file in list_of_files_in_folder:
    full_path = folder_location + "\\" + file
    print(full_path)
    list_of_full_paths.append(full_path)
    
print(list_of_full_paths)
# loop through list of full paths and select files which are for excel files
for file in list_of_full_paths:
    if "xlsx" in file:
        #print(file)

        # open excel files
        excel_wb_open = openpyxl.load_workbook(file)
            # search through all tables
        for sheet in excel_wb_open:
            #print(sheet.title)
            # search through all cells in a sheet
            for row in sheet.iter_rows():
                #print(row)
                    # check if search tearm occurs in a cell
                    # if search term is found in cell, then print file name, table name, cell location
                for cell in row:
                    if search_term in str(cell.value):
                        print("****Match found ****")
                        print(file, sheet, cell.coordinate)
                        print("****")
                        
                    


    

import pandas as pd
import openpyxl as xl
import os
import sys

##open excel file to write to
book = xl.Workbook()
sheet = book.active

## keywords to look for
search1 = "Max Proximal 1 Force"
search2 = "Average Proximal 1 Force"


##append headers
sheet.append((search1, search2))


##for every file in the directory
for file in os.listdir(os.getcwd()):

    ##if the file is of the correct type (excluding the output file)
    if file.endswith('.xlsx') and "~$" not in file and "output" not in file:
        workbook = xl.load_workbook(filename=file)

        active_ws = workbook.active


        max_force_col = 0
        avg_force_col = 0
        is_found = False
        to_break = False
        max_force = sys.maxsize
        avg_force = sys.maxsize

        ##for each row in the active sheet
        for row in active_ws.rows:
            ##for each cell in the row
            for cell in row:
                ##if cell value is a string then check for keyword
                if (isinstance(cell.value, str)):

                    ##if we find the max, note the column
                    if cell.value.find(search1) != -1:
                        max_force_col = cell.column
                        is_found = True
                        print(f'max force')
                        continue
                    
                    ##if we find the average, also note the column
                    elif cell.value.find(search2) != -1:
                        avg_force_col = cell.column
                        is_found = True
                        continue
                
                #we know the data points are on the row below the row where the keywords are found
                if (is_found):
                    if (cell.column == max_force_col):
                        max_force = cell.value
                    elif(cell.column == avg_force_col):
                        print(f'avg force')
                        avg_force = cell.value

                    ##we've found the values in this file, it's time to break out of it
                    to_break = True

            if(to_break):
                sheet.append((max_force, avg_force))
                break

book.save("output.xlsx")
                    


                



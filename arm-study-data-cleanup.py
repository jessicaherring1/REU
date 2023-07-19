### ORGANIZE EXCEL SHEET FOR QUALTRICS ARM MORPHOLOGY STUDY ###

from openpyxl import Workbook, load_workbook



# old workbook varibales 
oWorkbook = load_workbook(filename="Documents/REU/old-arm-study-data.xlsx")
oSheet = oWorkbook.active

# updated workbook varibales 
uWorkbook = Workbook()
uSheet = uWorkbook.active

#simplified workbook variables 
sWorkbook = Workbook()
sSheet = uWorkbook.active

# adds blank row to help with pandas processing
uSheet["A1"] = "1Title"

##CONVERT ALL COLUMNS INTO ROWS 

#for every columm
for i in range(1, oSheet.max_column):
    names = []

    #collect all information in column into an array
    for row in oSheet:
        name = row[i].value
        names.append(name)

    #add array into row
    for b in range(len(names)):
        temp = names[b]
        uSheet.cell(row=i+1, column=b+1).value = temp



## CLEAN OUT (delete rows that deal with not-arm specific data, timing, free-response answers, and price senstivity questions)

# outer for-loop ensures no lines are skipped as lines are deleted (very inefficient lol)
for i in range(1, uSheet.max_row):

    #for each row in the sheet
    for row in uSheet:

        #value in first column
        string = str(row[0].value) 

        # first letter of value in first column
        first_letter = string[0]

        #if first letter is not a number, delete the row
        if not first_letter.isnumeric():
            uSheet.delete_rows(row[0].row, 1)  

        #if stim, ps, or q49 are substrings within the cell value, delete the row
        if "stim" in string or "PS" in string or "Q49" in string:
            uSheet.delete_rows(row[0].row, 1)   



#delete columns B and C because they contain unnecessary data
uSheet.delete_cols(2, 2) 

# save into new excel file
uWorkbook.save(filename="Documents/REU/updated-arm-study-data.xlsx")

## SIMPLIFIED WORKBOOK

comp = ["1", "2", "3", "4", "5", "6"]
disc = ["7", "8", "9", "10", "11", "12"]


rows = uSheet.iter_rows(min_row=2, max_row=6, max_col=1, values_only=True)
values = [row[0] for row in rows]
print(values)
avg = sum(values) / len(values)

#for each row in the sheet
#for row in uSheet:

    #value in first column
#    string = str(row[0].value) 

    # first letter of value in first column
#    first_letter = string[0]

    # last letter of value in first column
#    last_2letter = string[::-1][0:2]

    #removes "11" and "12" questions
 #   if not last_2letter.isnumeric():
 #       last_letter = string[-1]
        #only gets questions surrounding competency
#        if "comp" in string and any([x in last_letter for x in comp]):
 #           for i in range(uSheet.max_column):
 #               sSheet.add(row)
                


## CREATING A CSV FILE
#https://www.studytonight.com/post/converting-xlsx-file-to-csv-file-using-python#:~:text=You%20can%20use%20openpyxl%20to,standard%20file%20I%2FO%20operations.
csv = open("Documents/REU/updated-arm-study-data.csv", "w+")

for row in uSheet.rows:
    l = list(row)
    for i in range(len(l)):
        if i == len(l) - 1:
            csv.write(str(l[i].value))
            csv.write('\n')
        else:
            csv.write(str(l[i].value) + ',')
        
csv.close()



'''
### NOTES ###

#ADD INDIVIDUAL VALUES TO CELLS 
uSheet["A1"] = "hello"
uSheet["B1"] = "world"

#STORE VALUE OF COLUMN INTO ARRAY
names = []
for row in oSheet:
    name = row[30].value
    names.append(name)

#STORE VALUE OF ARRAY INTO ROW
for b in range(len(names)):
        temp = names[b]
        uSheet.cell(row=1, column=b+1).value = temp
'''
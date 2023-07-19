### ORGANIZE EXCEL SHEET FOR QUALTRICS ARM MORPHOLOGY STUDY ###

from openpyxl import Workbook, load_workbook



# old workbook varibales 
oWorkbook = load_workbook(filename="Documents/REU/old-arm-study-data.xlsx")
oSheet = oWorkbook.active

# updated workbook (formatted how i like) 
uWorkbook = Workbook()
uSheet = uWorkbook.active

#simplified workbook (updated from uWorkbook, averages all participant answers per row) 
sWorkbook = Workbook()
sSheet = sWorkbook.active

# adds "blank" row to help with pandas processing
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

num_col = uSheet.max_column

#for each row in the sheet
for i in range (1, (uSheet.max_row + 1)):
    rows = uSheet.iter_cols(min_row=i, max_col=1, values_only=True)
    values = [row[0] for row in rows]

    print(values[0])

    location = "A" + str(i)
    print(location)
    sSheet[location] = values[0]


#for each row in the sheet
for i in range (2, (uSheet.max_row + 1)):
    rows = uSheet.iter_cols(min_row=i, min_col=2, max_col=num_col, values_only=True)
    values = [row[0] for row in rows]

    avg = sum(values) / len(values)

    location = "B" + str(i)

    sSheet[location] = avg

sSheet["B1"] = "Average"


# save into new excel file
sWorkbook.save(filename="Documents/REU/simplified-arm-study-data.xlsx")

                


## CREATING CSV FILES
#https://www.studytonight.com/post/converting-xlsx-file-to-csv-file-using-python#:~:text=You%20can%20use%20openpyxl%20to,standard%20file%20I%2FO%20operations.
uCsv = open("Documents/REU/updated-arm-study-data.csv", "w+")
sCsv = open("Documents/REU/simplified-arm-study-data.csv", "w+")

#for updated sheet
for row in uSheet.rows:
    l = list(row)
    for i in range(len(l)):
        if i == len(l) - 1:
            uCsv.write(str(l[i].value))
            uCsv.write('\n')
        else:
            uCsv.write(str(l[i].value) + ',')

#for simplified sheet
for row in sSheet.rows:
    l = list(row)
    for i in range(len(l)):
        if i == len(l) - 1:
            sCsv.write(str(l[i].value))
            sCsv.write('\n')
        else:
            sCsv.write(str(l[i].value) + ',')
        
uCsv.close()
sCsv.close()



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


#EXTRA INCASE NEED LATER

#for each row in the sheet
for row in uSheet:

    #value in first column
    string = str(row[0].value) 

    # first letter of value in first column
    first_letter = string[0]

    # last letter of value in first column
    last_2letter = string[::-1][0:2]

    #removes "11" and "12" questions
    if not last_2letter.isnumeric():
        last_letter = string[-1]
        #only gets questions surrounding competency
        if "comp" in string and any([x in last_letter for x in comp]):
            for i in range(uSheet.max_column):
                sSheet.add(row)


#Comp = ["1", "2", "3", "4", "5", "6"]
#disc = ["7", "8", "9", "10", "11", "12"]
'''
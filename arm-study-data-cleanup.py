### ORGANIZE EXCEL SHEET FOR QUALTRICS ARM MORPHOLOGY STUDY ###

from openpyxl import Workbook, load_workbook

# old workbook varibales 
oWorkbook = load_workbook(filename="Documents/REU/old-arm-study-data.xlsx")
oSheet = oWorkbook.active

# updated workbook (formatted how i like) 
uWorkbook = Workbook()
uSheet = uWorkbook.active

#averaged workbook (updated from uWorkbook, averages all participant answers per row) 
aWorkbook = Workbook()
aSheet = aWorkbook.active

#simplified workbook (updated from aWorkbook, simplifies questions into their categories (competency, discomfort, etc.)) 
sWorkbook = Workbook()
sSheet = sWorkbook.active

def main():


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


    ## AVERAGED WORKBOOK

    num_col = uSheet.max_column

    #copy over names from uSheet first column to aSheet first column
    for i in range (1, (uSheet.max_row + 1)):
        rows = uSheet.iter_cols(min_row=i, max_col=1, values_only=True)
        
        #all titlese 
        values = [row[0] for row in rows]

        location = "A" + str(i)
        
        aSheet[location] = values[0]


    #averaged all participant responses and copied them into second column
    for i in range (2, (uSheet.max_row + 1)):
        rows = uSheet.iter_cols(min_row=i, min_col=2, max_col=num_col, values_only=True)
        
        #all participant values 
        values = [row[0] for row in rows]

        avg = sum(values) / len(values)

        location = "B" + str(i)

        aSheet[location] = avg

    #column title
    aSheet["B1"] = "Average"

    # save into new excel file
    aWorkbook.save(filename="Documents/REU/averaged-arm-study-data.xlsx")



    ## SIMPLIFIED WORKBOOK
    comp = ["1", "2", "3", "4", "5", "6"]
    disc = ["7", "8", "9", "10", "11", "12"]

    disc1 = 0
    comp1 = 0
    saf1 = 0

    for i in range (1, (aSheet.max_row + 1)):
        rows = aSheet.iter_cols(min_row=i, min_col=1, max_col=2, values_only=True)

        #values in each column of the row
        values = [row[0] for row in rows]

        name = values[0]

        #number of the arm 
        first_letter = str(name)[0]

        last_2letter = name[::-1][0:2][::-1]

        if "1" in first_letter:
            value = values[1]
            if "comp" in name: 
                #discomfort questions 
                if last_2letter.isnumeric() or any([x in last_2letter for x in disc]):
                    disc1 += value
                #competency questions 
                else:
                    comp1 += value
            #safety questions
            elif "Average" not in str(value): 
                saf1 += value
                

    discomfort1 = disc1/6
    competency1 = comp1/6
    safety1 = saf1/3

    print(discomfort1, competency1, safety1)
                    

        



    #for each row in the sheet
    #for row in uSheet:

        #value in first column
    #   string = str(row[0].value) 

        # first letter of value in first column
    #   first_letter = string[0]

        # last letter of value in first column
    #   last_2letter = string[::-1][0:2]

        #removes "11" and "12" questions
    #   if not last_2letter.isnumeric():
    #       last_letter = string[-1]
            #only gets questions surrounding competency
    #       if "comp" in string and any([x in last_letter for x in comp]):
    #           for i in range(uSheet.max_column):
    #               sSheet.add(row)





    # save into new excel file
    sWorkbook.save(filename="Documents/REU/simplified-arm-study-data.xlsx")



    ## CREATING CSV FILES
    #https://www.studytonight.com/post/converting-xlsx-file-to-csv-file-using-python#:~:text=You%20can%20use%20openpyxl%20to,standard%20file%20I%2FO%20operations.
    uCsv = open("Documents/REU/updated-arm-study-data.csv", "w+")
    sCsv = open("Documents/REU/averaged-arm-study-data.csv", "w+")

    #for updated sheet
    for row in uSheet.rows:
        l = list(row)
        for i in range(len(l)):
            if i == len(l) - 1:
                uCsv.write(str(l[i].value))
                uCsv.write('\n')
            else:
                uCsv.write(str(l[i].value) + ',')

    #for averaged sheet
    for row in aSheet.rows:
        l = list(row)
        for i in range(len(l)):
            if i == len(l) - 1:
                sCsv.write(str(l[i].value))
                sCsv.write('\n')
            else:
                sCsv.write(str(l[i].value) + ',')
            
    uCsv.close()
    sCsv.close()

def simplify(x):
    comp = ["1", "2", "3", "4", "5", "6"]
    disc = ["7", "8", "9", "10", "11", "12"]

    disc = 0
    comp = 0
    saf = 0

    for i in range (1, (aSheet.max_row + 1)):
        rows = aSheet.iter_cols(min_row=i, min_col=1, max_col=2, values_only=True)

        #values in each column of the row
        values = [row[0] for row in rows]

        name = values[0]

        #number of the arm 
        first_letter = str(name)[0]

        last_2letter = name[::-1][0:2][::-1]

        if x in first_letter:
            value = values[1]
            if "comp" in name: 
                #discomfort questions 
                if last_2letter.isnumeric() or any([x in last_2letter for x in disc]):
                    disc += value
                #competency questions 
                else:
                    comp += value
            #safety questions
            elif "Average" not in str(value): 
                saf += value
                

    discomfort = disc/6
    competency = comp/6
    safety = saf/3

    return(discomfort, competency, safety)

main()


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
                aSheet.add(row)


#Comp = ["1", "2", "3", "4", "5", "6"]
#disc = ["7", "8", "9", "10", "11", "12"]
'''
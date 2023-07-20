### ORGANIZE EXCEL SHEET FOR QUALTRICS ARM MORPHOLOGY STUDY ###

from openpyxl import Workbook, load_workbook
import math

# old workbook varibales 
oWorkbook = load_workbook(filename="Documents/REU/old-arm-study-data.xlsx")
oSheet = oWorkbook.active

# updated workbook (formatted how i like) 
uWorkbook = Workbook()
uSheet = uWorkbook.active

#averaged-questions workbook (updated from uWorkbook, averages individual participant answers from question groups) 
qWorkbook = Workbook()
qSheet = qWorkbook.active

#averaged workbook (updated from uWorkbook, averages all participant answers per row) 
aWorkbook = Workbook()
aSheet = aWorkbook.active

#simplified workbook (updated from aWorkbook, simplifies questions into their categories (competency, discomfort, etc.)) 
sWorkbook = Workbook()
sSheet = sWorkbook.active


def main():

    uSheetFunc()

    qSheetFunc()

    simplify2("1", 2)

    aSheetFunc()

    sSheetFunc()

    csvFunc()



## UPDATES THE ORIGINAL EXCEL SHEET BY CONVERTING ALL COLUMNS INTO ROWS AND DELETE UNNECESSARY DATA
def uSheetFunc():

    # adds "blank"/title row to uSheet help with pandas processing
    uSheet["A1"] = "1Title"

    ## CONVERT (columns into rows)

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


def qSheetFunc():

    for i in range(1, 28):
        if i % 3 == 2:
            insert("A", i+1, str(math.ceil(i/3)) + "_Discomfort" ) 
        elif i % 3 == 1:
            insert("A", i+1, str(math.ceil(i/3)) + "_Competency" ) 
        elif i % 3 == 0:
            insert("A", i+1, str(math.ceil(i/3)) + "_Safety" ) 

    j= 1
    k = 2

    for i in range(2, uSheet.max_column):
        comp, disc, saf = simplify2(j, i)

        qSheet.cell(row=k, column=i).value = comp
        qSheet.cell(row=k+1, column=i).value = disc
        qSheet.cell(row=k+2, column=i).value = saf

        j+=1
        k+=3


    # save into new excel file
    qWorkbook.save(filename="Documents/REU/averages-questions-arm-study-data.xlsx")

## AVERAGED WORKBOOK -- AVERAGES ALL PARTIPANT VALUES AND PUTS INTO ONE COLUMN
def aSheetFunc():
    
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



## SIMPLIFIED WORKBOOK -- COMBINES QUESTION GROUPS INTO SINGLE ROW MY AVERAGING DATA (Q1-6 collapsed to "competency", Q7-12 to "discomfort", Q1-3 to "safety")
##          THEN ADDED ALL VALUES TO NEW FILE
def sSheetFunc():
    
    disc1, comp1, saf1 = simplify("1")
    disc2, comp2, saf2 = simplify("2")
    disc3, comp3, saf3 = simplify("3")
    disc4, comp4, saf4 = simplify("4")
    disc5, comp5, saf5 = simplify("5")
    disc6, comp6, saf6 = simplify("6")
    disc7, comp7, saf7 = simplify("7")
    disc8, comp8, saf8 = simplify("8")
    disc9, comp9, saf9 = simplify("9")

    for i in range(1, 28):
        if i % 3 == 2:
            insert("A", i+1, str(math.ceil(i/3)) + "_Discomfort" ) 
        elif i % 3 == 1:
            insert("A", i+1, str(math.ceil(i/3)) + "_Competency" ) 
        elif i % 3 == 0:
            insert("A", i+1, str(math.ceil(i/3)) + "_Safety" ) 

    insert("A", 1, "Title" ) 
    insert("B", 1, "Average") 

    insert("B", 2, disc1)
    insert("B", 3, comp1) 
    insert("B", 4, saf1) 
    insert("B", 5, disc2) 
    insert("B", 6, comp2) 
    insert("B", 7, saf2) 
    insert("B", 8, disc3) 
    insert("B", 9, comp3) 
    insert("B", 10, saf3) 
    insert("B", 11, disc4) 
    insert("B", 12, comp4)
    insert("B", 13, saf4) 
    insert("B", 14, disc5) 
    insert("B", 15, comp5) 
    insert("B", 16, saf5) 
    insert("B", 17, disc6) 
    insert("B", 18, comp6) 
    insert("B", 19, saf6)
    insert("B", 20, disc7)
    insert("B", 21, comp7) 
    insert("B", 22, saf7) 
    insert("B", 23, disc8) 
    insert("B", 24, comp8) 
    insert("B", 25, saf8)  
    insert("B", 26, disc9) 
    insert("B", 27, comp9) 
    insert("B", 28, saf9)  


    # save into new excel file
    sWorkbook.save(filename="Documents/REU/simplified-arm-study-data.xlsx")



#SIMPLIFIES THE QUESTION GROUPS AND RETURNS AVERAGE FOR EACH GROUP
def simplify(num):

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

        last_2letter = str(name)[::-1][0:2][::-1]

        if num in first_letter:
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
                

    #compute averages
    discomfort = disc1/6
    competency = comp1/6
    safety = saf1/3

    return(discomfort, competency, safety)



def simplify2(num, col):

    comp = ["1", "2", "3", "4", "5", "6"]
    disc = ["7", "8", "9", "10", "11", "12"]

    disc1 = 0
    comp1 = 0
    saf1 = 0


    for i in range (2, (uSheet.max_row + 1)):
        rows = uSheet.iter_cols(min_row=i,  max_col=col, values_only=True)

        #values in each column of the row
        values = [row[0] for row in rows]

        name = values[0]

        #number of the arm 
        first_letter = str(name)[0]

        last_2letter = name[::-1][0:2][::-1]

        if str(num) in first_letter:
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
                

    #compute averages
    discomfort = disc1/6
    competency = comp1/6
    safety = saf1/3

    return(discomfort, competency, safety)




#INSERTS INFO. INTO SPECIFIED LOCATION IN SSHEET
def insert(col, row, val):
    loc = col + str(row)
    sSheet[loc] = val



#SAVES EXCEL FILES IN CSV FORMAT
def csvFunc():
    ## CREATING CSV FILES
    #https://www.studytonight.com/post/converting-xlsx-file-to-csv-file-using-python#:~:text=You%20can%20use%20openpyxl%20to,standard%20file%20I%2FO%20operations.
    uCsv = open("Documents/REU/updated-arm-study-data.csv", "w+")
    aCsv = open("Documents/REU/averaged-arm-study-data.csv", "w+")
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

    #for averaged sheet
    for row in aSheet.rows:
        l = list(row)
        for i in range(len(l)):
            if i == len(l) - 1:
                aCsv.write(str(l[i].value))
                aCsv.write('\n')
            else:
                aCsv.write(str(l[i].value) + ',')

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
    aCsv.close()
    sCsv.close()


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
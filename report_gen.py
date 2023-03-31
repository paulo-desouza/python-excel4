from excel_methods import *

# TO-DO:
    # if not friday, aproximate to friday.



file_list = get_file_names()


# filter: only the YAML files shall pass. 
# use list comprehension


file_list = [file for file in file_list if file[-4:] == "yaml"]


# how many different names?


names_list = [file.split("_")[0] for file in file_list]

people = {}

for name in names_list:
    if name not in people:
        people[name] = []
        


# each value in the "people" list will be a column of data in the report ;
# aka a dictionary key for storing the YAMLs as a list in the value. 


for file in file_list:
    
    for name in people.keys():
        
        if name == file.split("_")[0]:
            people[name].append(file)


# decide time structure:
    
# start with it simple, and do it weekly:
    
    # - get earliest and latest weeks from the list
    # build out a table with each column being one (1) week, and 
    # each row being one person. 
    
# extract all the dates from file_list and sort THOSE. 

dates = []

for file in file_list: 
    dates.append(file.split("_")[1])

dates = [date_conversion(d) for d in dates]

dates = sorted(dates)

# produce a new "dates" list, containing all weeks in between the earliest and 
# latest weeks.

td = timedelta(days=7)

all_weeks = [dates[0]]


c = 0
while True:
    
    all_weeks.append(all_weeks[c]+td)
    c += 1

    
    if all_weeks[len(all_weeks)-1]+td > dates[len(dates)-1]:
        break


# now we can initiate the new excel sheet and position the table 
# according to the plan



wb = Workbook()

ws = wb.active



# write weekly columns:
    
col = 2
for week in all_weeks:
    char = get_column_letter(col)
    
    ws[char + "2"].value = week
    
    col += 1

row = 3
for name in people: 
    
    ws["A"+str(row)].value = name
    
    row += 1



# Match the dates of the reports with the dates on the columns. 
# If Present, process the found YAML file and output the data.
# Else; print "Data Unavailable"


for i, lst in enumerate(people.values()):
    
    for item in lst:  
        
        row = str(i + 3)
        col = 1
        
        for j in range(0, len(all_weeks)):
            
            col += 1
            char = get_column_letter(col)
            
            
            if date_conversion(item.split("_")[1]) == all_weeks[j]:
                # OUTPUT THE PROCESSED YAML DATA HERE!!!  
                ws[char+row].value = "test&success"
                
                
            elif ws[char+row].value != None:
                pass
                
            else:
                ws[char+row].value = "Data Unavailable :("
                
        


wb.save("output.xlsx")
























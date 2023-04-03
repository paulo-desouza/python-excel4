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

ws.title = "Celebree"
ws["A2"] = "Celebree"


#resize rows and columns

ws.column_dimensions["A"].width = 9
ws.row_dimensions[2].height = 50

for row in range(3, 50):
    for col in range(2, 50):
        char = get_column_letter(col)
        
        ws.column_dimensions[char].width = 38
        ws.row_dimensions[row].height = 170
        
for row in range(1, 50):
    for col in range(1, 50):
        char = get_column_letter(col)
        ws[char+str(row)].alignment = Alignment(horizontal = "center", vertical = "center", wrap_text=True)

# write weekly columns:
thick = Side(border_style="thin", color="000000")
col = 2
for week in all_weeks:
    char = get_column_letter(col)
    
    ws[char + "2"].value = week.strftime("Week of %B %d, %Y")
    ws[char + "2"].fill = PatternFill(fill_type='solid',
                                      start_color="99ff99",
                                      end_color='99ff99')
    ws[char + "2"].font = Font(bold = True)
    ws[char + "2"].border = Border(top=thick, left=thick, right=thick, bottom=thick)
    col += 1

row = 3
for name in people: 
    
    ws["A"+str(row)].value = name
    ws["A"+str(row)].fill = PatternFill(fill_type='solid',
                                      start_color='99ff99',
                                      end_color='99ff99')
    ws["A"+str(row)].border = Border(top=thick, left=thick, right=thick, bottom=thick)
    ws["A"+str(row)].font = Font(bold = True)
    row += 1



# Match the dates of the reports with the dates on the columns. 
# If Present, process the found YAML file and output the data.
# Else; print "Data Unavailable"


ws3 = wb.create_sheet('Data') 

for i, lst in enumerate(people.values()):
    
    for item in lst:  
        
        row = str(i + 3)
        col = 1
      
        for j in range(0, len(all_weeks)):
            
            
            col += 1
            char = get_column_letter(col)
            
            
            if date_conversion(item.split("_")[1]) == all_weeks[j]:
                
                
                
                with open(item, 'r') as file:
                    yaml_data = yaml.safe_load(file)
                    
                chart_data = [
                    
                    [(f"{yaml_data['EA']}_{yaml_data['Week']}", "Goals", "Actuals"),
                    
                    ("Inbound Calls", yaml_data['Table1']["GoalInbound"], yaml_data['Table1']["TotalInbound"]),
                    
                    ("Outbound Calls", yaml_data['Table1']["GoalOutbound"],  yaml_data['Table1']["TotalOutbound"]),
                    
                    ("Visits Scheduled", yaml_data['Table1']["GoalVisit"], yaml_data['Table1']["TotalVisit"]),
                     
                    ("Enrolled", yaml_data['Table1']["GoalEnrolled"], yaml_data['Table1']["TotalEnrolled"])],
                    
                    
                    
                    [(f"{yaml_data['EA']}_{yaml_data['Week']}", "Goals", "Actuals"),
                     
                    ("Inbound Calls", yaml_data['Table2']["GoalInbound"], yaml_data['Table2']["TotalInbound"]),
                     
                    ("Outbound Calls", yaml_data['Table2']["GoalOutbound"], yaml_data['Table2']["TotalOutbound"]),
                     
                    ("Links Sent", yaml_data['Table2']["GoalVisit"], yaml_data['Table2']["TotalSchedule"]),
                     
                    ("Enrolled", yaml_data['Table2']["GoalEnrolled"], yaml_data['Table2']["TotalEnrolled"])]
                ]
                
                for chart in chart_data:
                    for rw in chart:
                        ws3.append(rw)
            
                
                
                
                
counter = 1

for i, lst in enumerate(people.values()):
    
    for item in lst:  
        
        row = str(i + 3)
        col = 1
      
        for j in range(0, len(all_weeks)):
            
            
            col += 1
            char = get_column_letter(col)
            
            
            if date_conversion(item.split("_")[1]) == all_weeks[j]:
                
                chart = BarChart()
                chart.type = "col"
                chart.style = 10
                chart.height = 6
                chart.width = 7
                
                
                data = Reference(ws3, min_col=2, min_row=1, max_row=5, max_col=3)
                cats = Reference(ws3, min_col=1, min_row=2, max_row=5)
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(cats)
                chart.shape = 4
                ws.add_chart(chart, char+row)
                
                
                
                
                counter += 10         
                
                
                
                
            elif ws[char+row].value != None:
                pass
                
            else:
                ws[char+row].value = "Data Unavailable :("
                
            
                
                
ws2 = wb.create_sheet('Caliday', 1)  
ws2["A2"] = "Caliday"

#resize rows and columns

ws2.column_dimensions["A"].width = 9
ws2.row_dimensions[2].height = 50

for row in range(3, 50):
    for col in range(2, 50):
        char = get_column_letter(col)
        
        ws2.column_dimensions[char].width = 38
        ws2.row_dimensions[row].height = 170
        
for row in range(1, 50):
    for col in range(1, 50):
        char = get_column_letter(col)
        ws2[char+str(row)].alignment = Alignment(horizontal = "center", vertical = "center", wrap_text=True)


# write weekly columns:
    
col = 2
for week in all_weeks:
    char = get_column_letter(col)
    
    ws2[char + "2"].value = week.strftime("Week of %B %d, %Y")
    ws2[char + "2"].fill = PatternFill(fill_type='solid',
                                      start_color="66ccff",
                                      end_color='66ccff')
    ws2[char + "2"].font = Font(bold = True)
    ws2[char + "2"].border = Border(top=thick, left=thick, right=thick, bottom=thick)
    col += 1

row = 3
for name in people: 
    
    ws2["A"+str(row)].value = name
    ws2["A"+str(row)].fill = PatternFill(fill_type='solid',
                                      start_color='66ccff',
                                      end_color='66ccff')
    ws2["A"+str(row)].border = Border(top=thick, left=thick, right=thick, bottom=thick)
    ws2["A"+str(row)].font = Font(bold = True)
    row += 1


# Match the dates of the reports with the dates on the columns. 
# If Present, process the found YAML file and output the data.
# Else; print "Data Unavailable"


counter = 1

for i, lst in enumerate(people.values()):
    
    for item in lst:  
        
        row = str(i + 3)
        col = 1
      
        for j in range(0, len(all_weeks)):
            
            
            col += 1
            char = get_column_letter(col)
            
            
            if date_conversion(item.split("_")[1]) == all_weeks[j]:
                
                chart = BarChart()
                chart.type = "col"
                chart.style = 10
                chart.height = 6
                chart.width = 7
                
                
                data = Reference(ws3, min_col=2, min_row=6, max_row=10, max_col=3)
                cats = Reference(ws3, min_col=1, min_row=7, max_row=10)
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(cats)
                chart.shape = 4
                ws2.add_chart(chart, char+row)
                
                
                
                
                counter += 10         
                
                
                
                
            elif ws2[char+row].value != None:
                pass
                
            else:
                ws2[char+row].value = "Data Unavailable :("
                
                
                
                
ws.sheet_view.zoomScale = 110
ws2.sheet_view.zoomScale = 110
wb.save("output.xlsx")
























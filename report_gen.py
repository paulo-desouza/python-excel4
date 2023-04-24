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

ws.title = "OVERVIEW"



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


# Match the dates of the reports with the dates on the columns. 
# If Present, process the found YAML file and output the data.
# Else; print "Data Unavailable"




# create a list for each data row

ce_inbound_reported =  ["Inbound Calls"]
ce_outbound_reported = ["Outbound Calls"]
ce_visits_reported =   ["Visits Scheduled"]
ce_enrolled_reported = ["Enrolled"]
ce_fte_reported =      ["FTEs Enrolled"]

ce_inbound_goals =     ["Inbound Goals"]
ce_outbound_goals =    ["Outbound Goals"]
ce_visits_goals =      ["Visits Scheduled Goals"]
ce_enrolled_goals =    ["Enrolled Goals"]
ce_fte_goals =         ["FTEs Goals"]


ca_inbound_reported =  ["Inbound Calls"]
ca_outbound_reported = ["Outbound Calls"]
ca_links_reported =   ["Links Sent"]
ca_enrolled_reported =    ["Enrolled"]

ca_inbound_goals =     ["Inbound Goals"]
ca_outbound_goals =    ["Outbound Goals"]
ca_links_goals =       ["Links Sent Goals"]
ca_enrolled_goals =    ["Enrolled Goals"]

masterlist = [ce_inbound_reported, ce_outbound_reported, ce_visits_reported, ce_enrolled_reported,
              ce_fte_reported, ce_inbound_goals, ce_outbound_goals, ce_visits_goals, ce_enrolled_goals,
              ce_fte_goals, ca_inbound_reported, ca_outbound_reported, ca_links_reported, 
              ca_enrolled_reported, ca_inbound_goals, ca_outbound_goals, ca_links_goals, 
              ca_enrolled_goals]



people_dates = []

for i, f in enumerate(people.values()):
    
    people_dates.append([])
    
    for d in f:
        
        people_dates[i].append(str(date_conversion(d.split("_")[1])))
        
        

people_keys = list(people.keys())


for j in range(0, len(all_weeks)):
    
   
    
    for i, lst in enumerate(people.values()):
        j_check = True
    
        for item in lst:  
        
            with open(item, 'r') as file:
                    
                
                yaml_data = yaml.safe_load(file)
                
             
                
                
                if date_conversion(item.split("_")[1]) == all_weeks[j]:
                    
                    
                    ce_inbound_reported.append((yaml_data["EA"],
                                                yaml_data["Week"],
                                                yaml_data['Table1']["TotalInbound"]))
                    
                    ce_outbound_reported.append((yaml_data["EA"],
                                                 yaml_data["Week"], 
                                                 yaml_data['Table1']["TotalOutbound"]))
                    
                    ce_visits_reported.append((yaml_data["EA"], 
                                               yaml_data["Week"], 
                                               yaml_data['Table1']["TotalVisit"]))
                     
                    ce_enrolled_reported.append((yaml_data["EA"],
                                                 yaml_data["Week"], 
                                                 yaml_data['Table1']["TotalEnrolled"]))
                    
                    ce_fte_reported.append((yaml_data["EA"], 
                                            yaml_data["Week"], 
                                            yaml_data['Table1']["FTEs Actual"]))
                    
                    ###
                    
                    ce_inbound_goals.append((yaml_data["EA"], 
                                             yaml_data["Week"], 
                                             yaml_data['Table1']["GoalInbound"]))
                    
                    ce_outbound_goals.append((yaml_data["EA"], 
                                              yaml_data["Week"], 
                                              yaml_data['Table1']["GoalOutbound"]))
                    
                    ce_visits_goals.append((yaml_data["EA"], 
                                            yaml_data["Week"], 
                                            yaml_data['Table1']["GoalVisit"]))
                     
                    ce_enrolled_goals.append((yaml_data["EA"], 
                                              yaml_data["Week"], 
                                              yaml_data['Table1']["GoalEnrolled"]))
                    
                    ce_fte_goals.append((yaml_data["EA"], 
                                         yaml_data["Week"], 
                                         yaml_data['Table1']["FTEs Goal"]))
                    
                    
                    # ------------ #
                     
                    ca_inbound_reported.append((yaml_data["EA"], 
                                                yaml_data["Week"], 
                                                yaml_data['Table2']["TotalInbound"]))
                     
                    ca_outbound_reported.append((yaml_data["EA"],
                                                 yaml_data["Week"], 
                                                 yaml_data['Table2']["TotalOutbound"]))
                     
                    ca_links_reported.append((yaml_data["EA"], 
                                              yaml_data["Week"], 
                                              yaml_data['Table2']["TotalSchedule"]))
                     
                    ca_enrolled_reported.append((yaml_data["EA"],
                                                 yaml_data["Week"], 
                                                 yaml_data['Table2']["TotalEnrolled"]))
                    
                    ###
                    
                    ca_inbound_goals.append((yaml_data["EA"], 
                                             yaml_data["Week"],
                                             yaml_data['Table2']["GoalInbound"]))
                     
                    ca_outbound_goals.append((yaml_data["EA"], 
                                              yaml_data["Week"],
                                              yaml_data['Table2']["GoalOutbound"]))
                     
                    ca_links_goals.append((yaml_data["EA"],
                                           yaml_data["Week"],
                                           yaml_data['Table2']["GoalVisit"]))
                     
                    ca_enrolled_goals.append((yaml_data["EA"], 
                                              yaml_data["Week"], 
                                              yaml_data['Table2']["GoalEnrolled"]))
                    
                    
                    
                elif str(all_weeks[j]) not in people_dates[i] and j_check == True:
                    
                    if people_keys[i] == yaml_data["EA"]:
                    
                        for data in masterlist:
                            
                            data.append([yaml_data["EA"], all_weeks[j], "Data Unavailable"])
                            
                            j_check = False
                    
                    


table_weeks = all_weeks[:]

table_weeks.insert(0, "C")

for name in people.keys():
    counter = 2
    
    
    ws = wb.create_sheet(name + " CELEBREE DATA") 
    #ws.sheet_state = 'hidden'
    ws.append(table_weeks)
    
    
        
    for i, blob in enumerate(ce_enrolled_goals):
        
        if i == 0:
            ws["A2"].value = ce_inbound_reported[0]
            ws["A3"].value = ce_outbound_reported[0]
            ws["A4"].value = ce_visits_reported[0]
            ws["A5"].value = ce_enrolled_reported[0]
            ws["A6"].value = ce_fte_reported[0]
            ws.append(table_weeks)
            ws["A8"].value = ce_inbound_reported[0]
            ws["A9"].value = ce_inbound_goals[0]
            ws.append(table_weeks)
            ws["A11"].value = ce_outbound_reported[0]
            ws["A12"].value = ce_outbound_goals[0]
            ws.append(table_weeks)
            ws["A14"].value = ce_visits_reported[0]
            ws["A15"].value = ce_visits_goals[0]
            ws.append(table_weeks)
            ws["A17"].value = ce_enrolled_reported[0]
            ws["A18"].value = ce_enrolled_goals[0]
            ws.append(table_weeks)
            ws["A20"].value = ce_fte_reported[0]
            ws["A21"].value = ce_fte_reported[0]
            
            
            
            
            for colu in range(1, len(table_weeks)+1):
                
                char = get_column_letter(colu)
                ws[char+"31"].value = table_weeks[colu-1]
                
                
            ws["A32"].value = ca_inbound_reported[0]
            ws["A33"].value = ca_outbound_reported[0]
            ws["A34"].value = ca_links_reported[0]
            ws["A35"].value = ca_enrolled_reported[0]
            
            ws.append(table_weeks)
            
            
            ws["A37"].value = ca_inbound_reported[0]
            ws["A38"].value = ca_inbound_goals[0]
            ws.append(table_weeks)
            
            ws["A40"].value = ca_outbound_reported[0]
            ws["A41"].value = ca_outbound_goals[0]
            ws.append(table_weeks)
            
            ws["A43"].value = ca_links_reported[0]
            ws["A44"].value = ca_links_goals[0]
            ws.append(table_weeks)
            
            ws["A46"].value = ca_enrolled_reported[0]
            ws["A47"].value = ca_enrolled_goals[0]
            
            
            
            
        elif i != 0:
            if blob[0] == name:
                char = get_column_letter(counter)
                
                
                ws[char+"2"].value = ce_inbound_reported[i][2]
                ws[char+"3"].value = ce_outbound_reported[i][2]
                ws[char+"4"].value = ce_visits_reported[i][2]
                ws[char+"5"].value = ce_enrolled_reported[i][2]
                ws[char+"6"].value = ce_fte_reported[i][2]
                
                ws[char+"8"].value = ce_inbound_reported[i][2]
                ws[char+"9"].value = ce_inbound_goals[i][2]
                
                ws[char+"11"].value = ce_outbound_reported[i][2]
                ws[char+"12"].value = ce_outbound_goals[i][2]
                
                ws[char+"14"].value = ce_visits_reported[i][2]
                ws[char+"15"].value = ce_visits_goals[i][2]
                
                ws[char+"17"].value = ce_enrolled_reported[i][2]
                ws[char+"18"].value = ce_enrolled_goals[i][2]
                
                ws[char+"20"].value = ce_fte_reported[i][2]
                ws[char+"21"].value = ce_fte_reported[i][2]



                
                ws[char+"32"].value = ca_inbound_reported[i][2]
                ws[char+"33"].value = ca_outbound_reported[i][2]
                ws[char+"34"].value = ca_links_reported[i][2]
                ws[char+"35"].value = ca_enrolled_reported[i][2]
                
                
                ws[char+"37"].value = ca_inbound_reported[i][2]
                ws[char+"38"].value = ca_inbound_goals[i][2]
                
                ws[char+"40"].value = ca_outbound_reported[i][2]
                ws[char+"41"].value = ca_outbound_goals[i][2]
                
                ws[char+"43"].value = ca_links_reported[i][2]
                ws[char+"44"].value = ca_links_goals[i][2]
                
                ws[char+"46"].value = ca_enrolled_reported[i][2]
                ws[char+"47"].value = ca_enrolled_goals[i][2]
                
                counter += 1





ws = wb.create_sheet("ALL CELEBREE DATA") 
#ws.sheet_state = 'hidden'
ws.append(table_weeks)



for data in masterlist:
    for i, d in enumerate(data):
    
        if i > 0 and d[2] == "Data Unavailable":
            d[2] = 0



for name in people.keys():
    counter = 2

    for i, blob in enumerate(ce_enrolled_goals):
        
        if i == 0:
            ws["A2"].value = ce_inbound_reported[0]
            ws["A3"].value = ce_outbound_reported[0]
            ws["A4"].value = ce_visits_reported[0]
            ws["A5"].value = ce_enrolled_reported[0]
            ws["A6"].value = ce_fte_reported[0]
            ws.append(table_weeks)
            ws["A8"].value = ce_inbound_reported[0]
            ws["A9"].value = ce_inbound_goals[0]
            ws.append(table_weeks)
            ws["A11"].value = ce_outbound_reported[0]
            ws["A12"].value = ce_outbound_goals[0]
            ws.append(table_weeks)
            ws["A14"].value = ce_visits_reported[0]
            ws["A15"].value = ce_visits_goals[0]
            ws.append(table_weeks)
            ws["A17"].value = ce_enrolled_reported[0]
            ws["A18"].value = ce_enrolled_goals[0]
            ws.append(table_weeks)
            ws["A20"].value = ce_fte_reported[0]
            ws["A21"].value = ce_fte_reported[0]
            
            
            for colu in range(1, len(table_weeks)+1):
                
                char = get_column_letter(colu)
                ws[char+"31"].value = table_weeks[colu-1]
                
                
            ws["A32"].value = ca_inbound_reported[0]
            ws["A33"].value = ca_outbound_reported[0]
            ws["A34"].value = ca_links_reported[0]
            ws["A35"].value = ca_enrolled_reported[0]
            
            ws.append(table_weeks)
            
            
            ws["A37"].value = ca_inbound_reported[0]
            ws["A38"].value = ca_inbound_goals[0]
            ws.append(table_weeks)
            
            ws["A40"].value = ca_outbound_reported[0]
            ws["A41"].value = ca_outbound_goals[0]
            ws.append(table_weeks)
            
            ws["A43"].value = ca_links_reported[0]
            ws["A44"].value = ca_links_goals[0]
            ws.append(table_weeks)
            
            ws["A46"].value = ca_enrolled_reported[0]
            ws["A47"].value = ca_enrolled_goals[0]
            
            
            
            
        else:
            
            if blob[0] == name:
            
                char = get_column_letter(counter)
                
                ws[char+"2"].value = 0
                ws[char+"3"].value = 0
                ws[char+"4"].value = 0
                ws[char+"5"].value = 0
                ws[char+"6"].value = 0
                
                ws[char+"8"].value = 0
                ws[char+"9"].value = 0
                
                ws[char+"11"].value = 0
                ws[char+"12"].value = 0
                
                ws[char+"14"].value = 0
                ws[char+"15"].value = 0
                
                ws[char+"17"].value = 0
                ws[char+"18"].value = 0
                
                ws[char+"20"].value = 0
                ws[char+"21"].value = 0
    
                ws[char+"32"].value = 0
                ws[char+"33"].value = 0
                ws[char+"34"].value = 0
                ws[char+"35"].value = 0
                
                
                ws[char+"37"].value = 0
                ws[char+"38"].value = 0
                
                ws[char+"40"].value = 0
                ws[char+"41"].value = 0
                
                ws[char+"43"].value = 0
                ws[char+"44"].value = 0
                
                ws[char+"46"].value = 0
                ws[char+"47"].value = 0
                
                counter += 1
            


for name in people.keys():
    counter = 2

    for i, blob in enumerate(ce_enrolled_goals):
        
        if i > 0 and blob[0] == name:
            
            char = get_column_letter(counter)
            

            
            ws[char+"2"].value += ce_inbound_reported[i][2]
            ws[char+"3"].value += ce_outbound_reported[i][2]
            ws[char+"4"].value += ce_visits_reported[i][2]
            ws[char+"5"].value += ce_enrolled_reported[i][2]
            ws[char+"6"].value += ce_fte_reported[i][2]
            
            ws[char+"8"].value += ce_inbound_reported[i][2]
            ws[char+"9"].value += ce_inbound_goals[i][2]
            
            ws[char+"11"].value += ce_outbound_reported[i][2]
            ws[char+"12"].value += ce_outbound_goals[i][2]
            
            ws[char+"14"].value += ce_visits_reported[i][2]
            ws[char+"15"].value += ce_visits_goals[i][2]
            
            ws[char+"17"].value += ce_enrolled_reported[i][2]
            ws[char+"18"].value += ce_enrolled_goals[i][2]
            
            ws[char+"20"].value += ce_fte_reported[i][2]
            ws[char+"21"].value += ce_fte_reported[i][2]



            
            ws[char+"32"].value += ca_inbound_reported[i][2]
            ws[char+"33"].value += ca_outbound_reported[i][2]
            ws[char+"34"].value += ca_links_reported[i][2]
            ws[char+"35"].value += ca_enrolled_reported[i][2]
            
            
            ws[char+"37"].value += ca_inbound_reported[i][2]
            ws[char+"38"].value += ca_inbound_goals[i][2]
            
            ws[char+"40"].value += ca_outbound_reported[i][2]
            ws[char+"41"].value += ca_outbound_goals[i][2]
            
            ws[char+"43"].value += ca_links_reported[i][2]
            ws[char+"44"].value += ca_links_goals[i][2]
            
            ws[char+"46"].value += ca_enrolled_reported[i][2]
            ws[char+"47"].value += ca_enrolled_goals[i][2]
        
        
            counter += 1
        
        
        
        
    
    
    
    
    
    
    
    


for name in people.keys(): 
    
    ws = wb.create_sheet(name)






                

wb.save("output.xlsx")
























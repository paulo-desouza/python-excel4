from excel_methods import *


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
    
    
# weekly or monthly check:
    
monthly = False    

for files in people.values():
    if len(files) > 12:
        monthly = True    




if monthly == False:

    # - get earliest and latest weeks from the list
    # build out a table with each column being one (1) week, and 
    # each row being one person. 
    
    
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
    
    
    # convert table weeks into readable format
    
    table_weeks = [week.strftime("%b, %d \n(Week %U)") for week in table_weeks]
        
    
    for name in people.keys(): 
        
        ws = wb.create_sheet(name)
        
    
    table_weeks.insert(0, "C")
    
    for name in people.keys():
        counter = 2
        
        ws = wb.create_sheet(name + " CELEBREE DATA") 
        ws.sheet_state = 'hidden'
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
                ws["A21"].value = ce_fte_goals[0]
                
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
    
            # GENERAL OVERVIEW ChART (CELEBREE)
                    
        c_all = BarChart()
        c_all.title = "Client Conversion Pipeline Chart (Celebree)"
        c_all.style = 2
        c_all.y_axis.title = 'CALLS'
        c_all.x_axis.title = 'TIMELINE'
    
        
    
        c_all.height = 21
        c_all.width = 29
    
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
        data = Reference(ws, min_col=1, min_row=2, max_col=len(table_weeks), max_row=6)   
        c_all.add_data(data, titles_from_data=True, from_rows=True)
        c_all.set_categories(labels)
                    
    
    
    
        # INBOUND OVERVIEW 
    
        c_inb = BarChart()
        c_inb.title = "Inbound Calls (Celebree)"
        c_inb.style = 3
        c_inb.y_axis.title = 'CALLS'
        c_inb.x_axis.title = 'TIMELINE'
    
        c_inb.height = 21
        c_inb.width = 29
    
    
    
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
        data = Reference(ws, min_col=1, min_row=8, max_col=len(table_weeks), max_row=9)   
        c_inb.add_data(data, titles_from_data=True, from_rows=True)
        c_inb.set_categories(labels)
    
    
    
    
    
    
    
        # OUTBOUND OVERVIEW 
    
        c_out = BarChart()
        c_out.title = "Outbound Calls (Celebree)"
        c_out.style = 4
        c_out.y_axis.title = 'CALLS'
        c_out.x_axis.title = 'TIMELINE'
    
    
        c_out.height = 21
        c_out.width = 29
    
    
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
        data = Reference(ws, min_col=1, min_row=11, max_col=len(table_weeks), max_row=12)   
        c_out.add_data(data, titles_from_data=True, from_rows=True)
        c_out.set_categories(labels)
                    
    
    
    
    
        # VISITS SCHEDULED OVERVIEW
    
        c_vis = BarChart()
        c_vis.title = "Visits Scheduled (Celebree)"
        c_vis.style = 5
        c_vis.y_axis.title = 'CALLS'
        c_vis.x_axis.title = 'TIMELINE'
    
        c_vis.height = 21
        c_vis.width = 29
    
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
        data = Reference(ws, min_col=1, min_row=14, max_col=len(table_weeks), max_row=15)   
        c_vis.add_data(data, titles_from_data=True, from_rows=True)
        c_vis.set_categories(labels)
                    
    
    
    
    
        # ENROLLED OVERVIEW
    
        c_enr = BarChart()
        c_enr.title = "Enrolled (Celebree)"
        c_enr.style = 6
        c_enr.y_axis.title = 'CALLS'
        c_enr.x_axis.title = 'TIMELINE'
    
        c_enr.height = 21
        c_enr.width = 29
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
        data = Reference(ws, min_col=1, min_row=17, max_col=len(table_weeks), max_row=18)   
        c_enr.add_data(data, titles_from_data=True, from_rows=True)
        c_enr.set_categories(labels)
                    
    
    
    
    
        # FTE OVERVIEW
    
        c_fte = BarChart()
        c_fte.title = "FTE Enrolled (Celebree)"
        c_fte.style = 7
        c_fte.y_axis.title = 'CALLS'
        c_fte.x_axis.title = 'TIMELINE'
    
        c_fte.height = 21
        c_fte.width = 29
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
        data = Reference(ws, min_col=1, min_row=20, max_col=len(table_weeks), max_row=21)   
        c_fte.add_data(data, titles_from_data=True, from_rows=True)
        c_fte.set_categories(labels)
    
    
    
    
        # GENERAL OVERVIEW ChART (Caliday)
                    
        ca_all = BarChart()
        ca_all.title = "Client Conversion Pipeline Chart (Caliday)"
        ca_all.style = 2
        ca_all.y_axis.title = 'CALLS'
        ca_all.x_axis.title = 'TIMELINE'
    
        ca_all.height = 21
        ca_all.width = 29
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
        data = Reference(ws, min_col=1, min_row=32, max_col=len(table_weeks), max_row=35)   
        ca_all.add_data(data, titles_from_data=True, from_rows=True)
        ca_all.set_categories(labels)
                    
    
    
    
        # INBOUND OVERVIEW 
    
        ca_inb = BarChart()
        ca_inb.title = "Inbound Calls (Caliday)"
        ca_inb.style = 3
        ca_inb.y_axis.title = 'CALLS'
        ca_inb.x_axis.title = 'TIMELINE'
    
        ca_inb.height = 21
        ca_inb.width = 29
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
        data = Reference(ws, min_col=1, min_row=37, max_col=len(table_weeks), max_row=38)   
        ca_inb.add_data(data, titles_from_data=True, from_rows=True)
        ca_inb.set_categories(labels)
    
    
    
    
    
    
    
        # OUTBOUND OVERVIEW 
    
        ca_out = BarChart()
        ca_out.title = "Outbound Calls (Caliday)"
        ca_out.style = 4
        ca_out.y_axis.title = 'CALLS'
        ca_out.x_axis.title = 'TIMELINE'
    
        ca_out.height = 21
        ca_out.width = 29
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
        data = Reference(ws, min_col=1, min_row=40, max_col=len(table_weeks), max_row=41)   
        ca_out.add_data(data, titles_from_data=True, from_rows=True)
        ca_out.set_categories(labels)
                    
    
    
    
    
        # VISITS SCHEDULED OVERVIEW
    
        ca_vis = BarChart()
        ca_vis.title = "Links Sent (Caliday)"
        ca_vis.style = 5
        ca_vis.y_axis.title = 'CALLS'
        ca_vis.x_axis.title = 'TIMELINE'
    
        ca_vis.height = 21
        ca_vis.width = 29
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
        data = Reference(ws, min_col=1, min_row=43, max_col=len(table_weeks), max_row=44)   
        ca_vis.add_data(data, titles_from_data=True, from_rows=True)
        ca_vis.set_categories(labels)
                    
    
    
    
    
        # ENROLLED OVERVIEW
    
        ca_enr = BarChart()
        ca_enr.title = "Enrolled (Caliday)"
        ca_enr.style = 6
        ca_enr.y_axis.title = 'CALLS'
        ca_enr.x_axis.title = 'TIMELINE'
    
        ca_enr.height = 21
        ca_enr.width = 29
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
        data = Reference(ws, min_col=1, min_row=46, max_col=len(table_weeks), max_row=47)   
        ca_enr.add_data(data, titles_from_data=True, from_rows=True)
        ca_enr.set_categories(labels)
    
    
        ws = wb[name] 
    
        ws.sheet_view.zoomScale = 80
    
    
    
        ws.add_chart(c_all, "a1")
        ws.add_chart(c_inb, "a41")
        ws.add_chart(c_out, "A81")
        ws.add_chart(c_vis, "a121")
        ws.add_chart(c_enr, "A161")
        ws.add_chart(c_fte, "a201")
    
        ws.add_chart(ca_all, "s1")
        ws.add_chart(ca_inb, "s41")
        ws.add_chart(ca_out, "s81")
        ws.add_chart(ca_vis, "s121")
        ws.add_chart(ca_enr, "s161")
    
    
    
    ws = wb.create_sheet("ALL CELEBREE DATA") 
    ws.sheet_state = 'hidden'
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
                
               
                
    # GENERAL OVERVIEW ChART (CELEBREE)
                
    c_all = LineChart()
    c_all.title = "Client Conversion Pipeline Chart (Celebree)"
    c_all.style = 2
    c_all.y_axis.title = 'CALLS'
    c_all.x_axis.title = 'TIMELINE'
    
    for s in c_all.series:
        s.smooth = True
    
    c_all.height = 21
    c_all.width = 29
    
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
    data = Reference(ws, min_col=1, min_row=2, max_col=len(table_weeks), max_row=6)   
    c_all.add_data(data, titles_from_data=True, from_rows=True)
    c_all.set_categories(labels)
                
    
    # INBOUND OVERVIEW 
    
    c_inb = LineChart()
    c_inb.title = "Inbound Calls (Celebree)"
    c_inb.style = 3
    c_inb.y_axis.title = 'CALLS'
    c_inb.x_axis.title = 'TIMELINE'
    
    c_inb.height = 21
    c_inb.width = 29
    
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
    data = Reference(ws, min_col=1, min_row=8, max_col=len(table_weeks), max_row=9)   
    c_inb.add_data(data, titles_from_data=True, from_rows=True)
    c_inb.set_categories(labels)
    
    
    # OUTBOUND OVERVIEW 
    
    c_out = LineChart()
    c_out.title = "Outbound Calls (Celebree)"
    c_out.style = 4
    c_out.y_axis.title = 'CALLS'
    c_out.x_axis.title = 'TIMELINE'
    
    
    c_out.height = 21
    c_out.width = 29
    
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
    data = Reference(ws, min_col=1, min_row=11, max_col=len(table_weeks), max_row=12)   
    c_out.add_data(data, titles_from_data=True, from_rows=True)
    c_out.set_categories(labels)
                
    
    # VISITS SCHEDULED OVERVIEW
    
    c_vis = LineChart()
    c_vis.title = "Visits Scheduled (Celebree)"
    c_vis.style = 5
    c_vis.y_axis.title = 'CALLS'
    c_vis.x_axis.title = 'TIMELINE'
    
    c_vis.height = 21
    c_vis.width = 29
    
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
    data = Reference(ws, min_col=1, min_row=14, max_col=len(table_weeks), max_row=15)   
    c_vis.add_data(data, titles_from_data=True, from_rows=True)
    c_vis.set_categories(labels)
                
    
    # ENROLLED OVERVIEW
    
    c_enr = LineChart()
    c_enr.title = "Enrolled (Celebree)"
    c_enr.style = 6
    c_enr.y_axis.title = 'CALLS'
    c_enr.x_axis.title = 'TIMELINE'
    
    c_enr.height = 21
    c_enr.width = 29
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
    data = Reference(ws, min_col=1, min_row=17, max_col=len(table_weeks), max_row=18)   
    c_enr.add_data(data, titles_from_data=True, from_rows=True)
    c_enr.set_categories(labels)
                
    
    # FTE OVERVIEW
    
    c_fte = LineChart()
    c_fte.title = "FTE Enrolled (Celebree)"
    c_fte.style = 7
    c_fte.y_axis.title = 'CALLS'
    c_fte.x_axis.title = 'TIMELINE'
    
    c_fte.height = 21
    c_fte.width = 29
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
    data = Reference(ws, min_col=1, min_row=20, max_col=len(table_weeks), max_row=21)   
    c_fte.add_data(data, titles_from_data=True, from_rows=True)
    c_fte.set_categories(labels)
    
    
    # GENERAL OVERVIEW ChART (Caliday)
                
    ca_all = LineChart()
    ca_all.title = "Client Conversion Pipeline Chart (Caliday)"
    ca_all.style = 2
    ca_all.y_axis.title = 'CALLS'
    ca_all.x_axis.title = 'TIMELINE'
    
    ca_all.height = 21
    ca_all.width = 29
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
    data = Reference(ws, min_col=1, min_row=32, max_col=len(table_weeks), max_row=35)   
    ca_all.add_data(data, titles_from_data=True, from_rows=True)
    ca_all.set_categories(labels)
    
    
    # INBOUND OVERVIEW 
    
    ca_inb = LineChart()
    ca_inb.title = "Inbound Calls (Caliday)"
    ca_inb.style = 3
    ca_inb.y_axis.title = 'CALLS'
    ca_inb.x_axis.title = 'TIMELINE'
    
    ca_inb.height = 21
    ca_inb.width = 29
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
    data = Reference(ws, min_col=1, min_row=37, max_col=len(table_weeks), max_row=38)   
    ca_inb.add_data(data, titles_from_data=True, from_rows=True)
    ca_inb.set_categories(labels)
    
    
    # OUTBOUND OVERVIEW 
    
    ca_out = LineChart()
    ca_out.title = "Outbound Calls (Caliday)"
    ca_out.style = 4
    ca_out.y_axis.title = 'CALLS'
    ca_out.x_axis.title = 'TIMELINE'
    
    ca_out.height = 21
    ca_out.width = 29
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
    data = Reference(ws, min_col=1, min_row=40, max_col=len(table_weeks), max_row=41)   
    ca_out.add_data(data, titles_from_data=True, from_rows=True)
    ca_out.set_categories(labels)
                
    
    # VISITS SCHEDULED OVERVIEW
    
    ca_vis = LineChart()
    ca_vis.title = "Links Sent (Caliday)"
    ca_vis.style = 5
    ca_vis.y_axis.title = 'CALLS'
    ca_vis.x_axis.title = 'TIMELINE'
    
    ca_vis.height = 21
    ca_vis.width = 29
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
    data = Reference(ws, min_col=1, min_row=43, max_col=len(table_weeks), max_row=44)   
    ca_vis.add_data(data, titles_from_data=True, from_rows=True)
    ca_vis.set_categories(labels)
                
    
    # ENROLLED OVERVIEW
    
    ca_enr = LineChart()
    ca_enr.title = "Enrolled (Caliday)"
    ca_enr.style = 6
    ca_enr.y_axis.title = 'CALLS'
    ca_enr.x_axis.title = 'TIMELINE'
    
    ca_enr.height = 21
    ca_enr.width = 29
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_weeks)) 
    data = Reference(ws, min_col=1, min_row=46, max_col=len(table_weeks), max_row=47)   
    ca_enr.add_data(data, titles_from_data=True, from_rows=True)
    ca_enr.set_categories(labels)
    
    
    ws = wb["OVERVIEW"]
    ws.sheet_view.zoomScale = 80
    
    
    ws.add_chart(c_all, "a1")
    ws.add_chart(c_inb, "a41")
    ws.add_chart(c_out, "A81")
    ws.add_chart(c_vis, "a121")
    ws.add_chart(c_enr, "A161")
    ws.add_chart(c_fte, "a201")
    
    ws.add_chart(ca_all, "s1")
    ws.add_chart(ca_inb, "s41")
    ws.add_chart(ca_out, "s81")
    ws.add_chart(ca_vis, "s121")
    ws.add_chart(ca_enr, "s161")
    
    
    wb.save("output1.xlsx")
    

else:
    
    all_months = {}
    monthly_files = {}
    
    
    # we need to now convert the code into monthly stuff. 
    
    dates = []
    
    for file in file_list: 
        dates.append(file.split("_")[1])
    
    
    dates = [date_conversion(d) for d in dates]
    
    
    for month in dates:
        
        month_list = list(all_months.keys())
        
        if month.strftime("%B") not in month_list:
            all_months[month.strftime("%B")] = [month]
            monthly_files[month.strftime("%B")] = []
            
        elif month.strftime("%B") in month_list:
            all_months[month.strftime("%B")].append(month)
            
    month_list = list(all_months.keys())
            
        
        
        
    for file in file_list:
    
        if date_conversion(file.split("_")[1]).strftime("%B") in list(all_months.keys()):
            
            monthly_files[date_conversion(file.split("_")[1]).strftime("%B")].append(file)
        
    # now we can initiate the new excel sheet and position the table 
    
    wb = Workbook()
    
    ws = wb.active
    
    ws.title = "OVERVIEW"
    
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
    ca_links_reported =    ["Links Sent"]
    ca_enrolled_reported = ["Enrolled"]
    
    ca_inbound_goals =     ["Inbound Goals"]
    ca_outbound_goals =    ["Outbound Goals"]
    ca_links_goals =       ["Links Sent Goals"]
    ca_enrolled_goals =    ["Enrolled Goals"]
    
    masterlist = [ce_inbound_reported, ce_outbound_reported, ce_visits_reported, ce_enrolled_reported,
                  ce_fte_reported, ce_inbound_goals, ce_outbound_goals, ce_visits_goals, ce_enrolled_goals,
                  ce_fte_goals, ca_inbound_reported, ca_outbound_reported, ca_links_reported, 
                  ca_enrolled_reported, ca_inbound_goals, ca_outbound_goals, ca_links_goals, 
                  ca_enrolled_goals]
    
    
    # list of first weeks
    the_list = []
    
    for month in all_months.values():
        
        the_list.append(month[0])
        
    
    
    
    
    
    cc = 1
    for name in people.keys():
        
        for month_name, month in monthly_files.items():
            
            for i, week in enumerate(month):
                
                with open(week, 'r') as file:
                    
                    yaml_data = yaml.safe_load(file)
                    
                    
                    if date_conversion(week.split("_")[1].split(".")[0]) in the_list and name == yaml_data["EA"]:
                
                            
                        ce_inbound_reported.append([yaml_data["EA"],
                                                    month_name,
                                                    yaml_data['Table1']["TotalInbound"]])
                        
                        ce_outbound_reported.append([yaml_data["EA"],
                                                     month_name, 
                                                     yaml_data['Table1']["TotalOutbound"]])
                        
                        ce_visits_reported.append([yaml_data["EA"], 
                                                   month_name, 
                                                   yaml_data['Table1']["TotalVisit"]])
                         
                        ce_enrolled_reported.append([yaml_data["EA"],
                                                     month_name, 
                                                     yaml_data['Table1']["TotalEnrolled"]])
                        
                        ce_fte_reported.append([yaml_data["EA"], 
                                                month_name, 
                                                yaml_data['Table1']["FTEs Actual"]])
                        
                        ###
                        
                        ce_inbound_goals.append([yaml_data["EA"], 
                                                 month_name, 
                                                 yaml_data['Table1']["GoalInbound"]])
                        
                        ce_outbound_goals.append([yaml_data["EA"], 
                                                  month_name, 
                                                  yaml_data['Table1']["GoalOutbound"]])
                        
                        ce_visits_goals.append([yaml_data["EA"], 
                                                month_name, 
                                                yaml_data['Table1']["GoalVisit"]])
                         
                        ce_enrolled_goals.append([yaml_data["EA"], 
                                                  month_name, 
                                                  yaml_data['Table1']["GoalEnrolled"]])
                        
                        ce_fte_goals.append([yaml_data["EA"], 
                                             month_name, 
                                             yaml_data['Table1']["FTEs Goal"]])
                        
                        
                        # ------------ #
                         
                        ca_inbound_reported.append([yaml_data["EA"], 
                                                    month_name, 
                                                    yaml_data['Table2']["TotalInbound"]])
                         
                        ca_outbound_reported.append([yaml_data["EA"],
                                                     month_name, 
                                                     yaml_data['Table2']["TotalOutbound"]])
                         
                        ca_links_reported.append([yaml_data["EA"], 
                                                  month_name, 
                                                  yaml_data['Table2']["TotalSchedule"]])
                         
                        ca_enrolled_reported.append([yaml_data["EA"],
                                                     month_name, 
                                                     yaml_data['Table2']["TotalEnrolled"]])
                        
                        ###
                        
                        ca_inbound_goals.append([yaml_data["EA"], 
                                                 month_name,
                                                 yaml_data['Table2']["GoalInbound"]])
                         
                        ca_outbound_goals.append([yaml_data["EA"], 
                                                  month_name,
                                                  yaml_data['Table2']["GoalOutbound"]])
                         
                        ca_links_goals.append([yaml_data["EA"],
                                               month_name,
                                               yaml_data['Table2']["GoalVisit"]])
                         
                        ca_enrolled_goals.append([yaml_data["EA"], 
                                                  month_name, 
                                                  yaml_data['Table2']["GoalEnrolled"]])
                        
                        
                        
                    elif name == yaml_data["EA"]:
                        
                        
                        
                        ce_inbound_reported[cc][2] += yaml_data['Table1']["TotalInbound"]
                        
                        ce_outbound_reported[cc][2] += yaml_data['Table1']["TotalOutbound"]
                        
                        ce_visits_reported[cc][2] += yaml_data['Table1']["TotalVisit"]
                         
                        ce_enrolled_reported[cc][2] += yaml_data['Table1']["TotalEnrolled"]
                        
                        ce_fte_reported[cc][2] += yaml_data['Table1']["FTEs Actual"]
                        
                        ce_inbound_goals[cc][2] += yaml_data['Table1']["GoalInbound"]
                        
                        ce_outbound_goals[cc][2] += yaml_data['Table1']["GoalOutbound"]
                        
                        ce_visits_goals[cc][2] += yaml_data['Table1']["GoalVisit"]
                         
                        ce_enrolled_goals[cc][2] += yaml_data['Table1']["GoalEnrolled"]
                        
                        ce_fte_goals[cc][2] += yaml_data['Table1']["FTEs Goal"]
                        
                        ca_inbound_reported[cc][2] += yaml_data['Table2']["TotalInbound"]
                         
                        ca_outbound_reported[cc][2] += yaml_data['Table2']["TotalOutbound"]
                         
                        ca_links_reported[cc][2] += yaml_data['Table2']["TotalSchedule"]
                         
                        ca_enrolled_reported[cc][2] += yaml_data['Table2']["TotalEnrolled"]
                        
                        ca_inbound_goals[cc][2] += yaml_data['Table2']["GoalInbound"]
                         
                        ca_outbound_goals[cc][2] += yaml_data['Table2']["GoalOutbound"]
                         
                        ca_links_goals[cc][2] += yaml_data['Table2']["GoalVisit"]
                         
                        ca_enrolled_goals[cc][2] += yaml_data['Table2']["GoalEnrolled"]
                        
                    
            cc += 1
                
        
        
    table_months = month_list[:]
    
    
    for name in people.keys(): 
        
        ws = wb.create_sheet(name)
        
    
    table_months.insert(0, "C")
    
    for name in people.keys():
        counter = 2
        
        ws = wb.create_sheet(name + " CELEBREE DATA") 
        ws.sheet_state = 'hidden'
        ws.append(table_months)
        
        for i, blob in enumerate(ce_enrolled_goals):
            
            if i == 0:
                
                ws["A2"].value = ce_inbound_reported[0]
                ws["A3"].value = ce_outbound_reported[0]
                ws["A4"].value = ce_visits_reported[0]
                ws["A5"].value = ce_enrolled_reported[0]
                ws["A6"].value = ce_fte_reported[0]
                
                ws.append(table_months)
                
                ws["A8"].value = ce_inbound_reported[0]
                ws["A9"].value = ce_inbound_goals[0]
                
                ws.append(table_months)
                
                ws["A11"].value = ce_outbound_reported[0]
                ws["A12"].value = ce_outbound_goals[0]
                
                ws.append(table_months)
                
                ws["A14"].value = ce_visits_reported[0]
                ws["A15"].value = ce_visits_goals[0]
                
                ws.append(table_months)
                
                ws["A17"].value = ce_enrolled_reported[0]
                ws["A18"].value = ce_enrolled_goals[0]
                
                ws.append(table_months)
                
                ws["A20"].value = ce_fte_reported[0]
                ws["A21"].value = ce_fte_goals[0]
                
                for colu in range(1, len(table_months)+1):
                    
                    char = get_column_letter(colu)
                    ws[char+"31"].value = table_months[colu-1]
                    
                ws["A32"].value = ca_inbound_reported[0]
                ws["A33"].value = ca_outbound_reported[0]
                ws["A34"].value = ca_links_reported[0]
                ws["A35"].value = ca_enrolled_reported[0]
                
                ws.append(table_months)
                
                ws["A37"].value = ca_inbound_reported[0]
                ws["A38"].value = ca_inbound_goals[0]
                
                ws.append(table_months)
                
                ws["A40"].value = ca_outbound_reported[0]
                ws["A41"].value = ca_outbound_goals[0]
                
                ws.append(table_months)
                
                ws["A43"].value = ca_links_reported[0]
                ws["A44"].value = ca_links_goals[0]
                
                ws.append(table_months)
                
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
        
             # GENERAL OVERVIEW ChART (CELEBREE)
                    
        c_all = BarChart()
        c_all.title = "Client Conversion Pipeline Chart (Celebree)"
        c_all.style = 2
        c_all.y_axis.title = 'CALLS'
        c_all.x_axis.title = 'TIMELINE'
    
        
    
        c_all.height = 21
        c_all.width = 29
    
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
        data = Reference(ws, min_col=1, min_row=2, max_col=len(table_months), max_row=6)   
        c_all.add_data(data, titles_from_data=True, from_rows=True)
        c_all.set_categories(labels)
                    
    
    
    
        # INBOUND OVERVIEW 
    
        c_inb = BarChart()
        c_inb.title = "Inbound Calls (Celebree)"
        c_inb.style = 3
        c_inb.y_axis.title = 'CALLS'
        c_inb.x_axis.title = 'TIMELINE'
    
        c_inb.height = 21
        c_inb.width = 29
    
    
    
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
        data = Reference(ws, min_col=1, min_row=8, max_col=len(table_months), max_row=9)   
        c_inb.add_data(data, titles_from_data=True, from_rows=True)
        c_inb.set_categories(labels)
    
    
    
    
    
    
    
        # OUTBOUND OVERVIEW 
    
        c_out = BarChart()
        c_out.title = "Outbound Calls (Celebree)"
        c_out.style = 4
        c_out.y_axis.title = 'CALLS'
        c_out.x_axis.title = 'TIMELINE'
    
    
        c_out.height = 21
        c_out.width = 29
    
    
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
        data = Reference(ws, min_col=1, min_row=11, max_col=len(table_months), max_row=12)   
        c_out.add_data(data, titles_from_data=True, from_rows=True)
        c_out.set_categories(labels)
                    
    
    
    
    
        # VISITS SCHEDULED OVERVIEW
    
        c_vis = BarChart()
        c_vis.title = "Visits Scheduled (Celebree)"
        c_vis.style = 5
        c_vis.y_axis.title = 'CALLS'
        c_vis.x_axis.title = 'TIMELINE'
    
        c_vis.height = 21
        c_vis.width = 29
    
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
        data = Reference(ws, min_col=1, min_row=14, max_col=len(table_months), max_row=15)   
        c_vis.add_data(data, titles_from_data=True, from_rows=True)
        c_vis.set_categories(labels)
                    
    
    
    
    
        # ENROLLED OVERVIEW
    
        c_enr = BarChart()
        c_enr.title = "Enrolled (Celebree)"
        c_enr.style = 6
        c_enr.y_axis.title = 'CALLS'
        c_enr.x_axis.title = 'TIMELINE'
    
        c_enr.height = 21
        c_enr.width = 29
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
        data = Reference(ws, min_col=1, min_row=17, max_col=len(table_months), max_row=18)   
        c_enr.add_data(data, titles_from_data=True, from_rows=True)
        c_enr.set_categories(labels)
                    
    
    
    
    
        # FTE OVERVIEW
    
        c_fte = BarChart()
        c_fte.title = "FTE Enrolled (Celebree)"
        c_fte.style = 7
        c_fte.y_axis.title = 'CALLS'
        c_fte.x_axis.title = 'TIMELINE'
    
        c_fte.height = 21
        c_fte.width = 29
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
        data = Reference(ws, min_col=1, min_row=20, max_col=len(table_months), max_row=21)   
        c_fte.add_data(data, titles_from_data=True, from_rows=True)
        c_fte.set_categories(labels)
    
    
    
    
        # GENERAL OVERVIEW ChART (Caliday)
                    
        ca_all = BarChart()
        ca_all.title = "Client Conversion Pipeline Chart (Caliday)"
        ca_all.style = 2
        ca_all.y_axis.title = 'CALLS'
        ca_all.x_axis.title = 'TIMELINE'
    
        ca_all.height = 21
        ca_all.width = 29
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
        data = Reference(ws, min_col=1, min_row=32, max_col=len(table_months), max_row=35)   
        ca_all.add_data(data, titles_from_data=True, from_rows=True)
        ca_all.set_categories(labels)
                    
    
    
    
        # INBOUND OVERVIEW 
    
        ca_inb = BarChart()
        ca_inb.title = "Inbound Calls (Caliday)"
        ca_inb.style = 3
        ca_inb.y_axis.title = 'CALLS'
        ca_inb.x_axis.title = 'TIMELINE'
    
        ca_inb.height = 21
        ca_inb.width = 29
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
        data = Reference(ws, min_col=1, min_row=37, max_col=len(table_months), max_row=38)   
        ca_inb.add_data(data, titles_from_data=True, from_rows=True)
        ca_inb.set_categories(labels)
    
    
    
    
    
    
    
        # OUTBOUND OVERVIEW 
    
        ca_out = BarChart()
        ca_out.title = "Outbound Calls (Caliday)"
        ca_out.style = 4
        ca_out.y_axis.title = 'CALLS'
        ca_out.x_axis.title = 'TIMELINE'
    
        ca_out.height = 21
        ca_out.width = 29
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
        data = Reference(ws, min_col=1, min_row=40, max_col=len(table_months), max_row=41)   
        ca_out.add_data(data, titles_from_data=True, from_rows=True)
        ca_out.set_categories(labels)
                    
    
    
    
    
        # VISITS SCHEDULED OVERVIEW
    
        ca_vis = BarChart()
        ca_vis.title = "Links Sent (Caliday)"
        ca_vis.style = 5
        ca_vis.y_axis.title = 'CALLS'
        ca_vis.x_axis.title = 'TIMELINE'
    
        ca_vis.height = 21
        ca_vis.width = 29
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
        data = Reference(ws, min_col=1, min_row=43, max_col=len(table_months), max_row=44)   
        ca_vis.add_data(data, titles_from_data=True, from_rows=True)
        ca_vis.set_categories(labels)
                    
    
    
    
    
        # ENROLLED OVERVIEW
    
        ca_enr = BarChart()
        ca_enr.title = "Enrolled (Caliday)"
        ca_enr.style = 6
        ca_enr.y_axis.title = 'CALLS'
        ca_enr.x_axis.title = 'TIMELINE'
    
        ca_enr.height = 21
        ca_enr.width = 29
    
        labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
        data = Reference(ws, min_col=1, min_row=46, max_col=len(table_months), max_row=47)   
        ca_enr.add_data(data, titles_from_data=True, from_rows=True)
        ca_enr.set_categories(labels)
    
    
        ws = wb[name] 
    
        ws.sheet_view.zoomScale = 80
    
    
    
        ws.add_chart(c_all, "a1")
        ws.add_chart(c_inb, "a41")
        ws.add_chart(c_out, "A81")
        ws.add_chart(c_vis, "a121")
        ws.add_chart(c_enr, "A161")
        ws.add_chart(c_fte, "a201")
    
        ws.add_chart(ca_all, "s1")
        ws.add_chart(ca_inb, "s41")
        ws.add_chart(ca_out, "s81")
        ws.add_chart(ca_vis, "s121")
        ws.add_chart(ca_enr, "s161")
        


                    
 
     
    
    ws = wb.create_sheet("ALL CELEBREE DATA") 
    ws.sheet_state = 'hidden'
    ws.append(table_months)
    
    
    
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
                
                ws.append(table_months)
                
                ws["A8"].value = ce_inbound_reported[0]
                ws["A9"].value = ce_inbound_goals[0]
                
                ws.append(table_months)
                
                ws["A11"].value = ce_outbound_reported[0]
                ws["A12"].value = ce_outbound_goals[0]
                
                ws.append(table_months)
                
                ws["A14"].value = ce_visits_reported[0]
                ws["A15"].value = ce_visits_goals[0]
                
                ws.append(table_months)
                
                ws["A17"].value = ce_enrolled_reported[0]
                ws["A18"].value = ce_enrolled_goals[0]
                
                ws.append(table_months)
                
                ws["A20"].value = ce_fte_reported[0]
                ws["A21"].value = ce_fte_reported[0]
                
                
                for colu in range(1, len(table_months)+1):
                    
                    char = get_column_letter(colu)
                    ws[char+"31"].value = table_months[colu-1]
                    
                    
                ws["A32"].value = ca_inbound_reported[0]
                ws["A33"].value = ca_outbound_reported[0]
                ws["A34"].value = ca_links_reported[0]
                ws["A35"].value = ca_enrolled_reported[0]
                
                ws.append(table_months)
                
                ws["A37"].value = ca_inbound_reported[0]
                ws["A38"].value = ca_inbound_goals[0]
                
                ws.append(table_months)
                
                ws["A40"].value = ca_outbound_reported[0]
                ws["A41"].value = ca_outbound_goals[0]
                
                ws.append(table_months)
                
                ws["A43"].value = ca_links_reported[0]
                ws["A44"].value = ca_links_goals[0]
                
                ws.append(table_months)
                
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
                
               
                
    # GENERAL OVERVIEW ChART (CELEBREE)
                
    c_all = LineChart()
    c_all.title = "Client Conversion Pipeline Chart (Celebree)"
    c_all.style = 2
    c_all.y_axis.title = 'CALLS'
    c_all.x_axis.title = 'TIMELINE'
    
    for s in c_all.series:
        s.smooth = True
    
    c_all.height = 21
    c_all.width = 29
    
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
    data = Reference(ws, min_col=1, min_row=2, max_col=len(table_months), max_row=6)   
    c_all.add_data(data, titles_from_data=True, from_rows=True)
    c_all.set_categories(labels)
                
    
    # INBOUND OVERVIEW 
    
    c_inb = LineChart()
    c_inb.title = "Inbound Calls (Celebree)"
    c_inb.style = 3
    c_inb.y_axis.title = 'CALLS'
    c_inb.x_axis.title = 'TIMELINE'
    
    c_inb.height = 21
    c_inb.width = 29
    
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
    data = Reference(ws, min_col=1, min_row=8, max_col=len(table_months), max_row=9)   
    c_inb.add_data(data, titles_from_data=True, from_rows=True)
    c_inb.set_categories(labels)
    
    
    # OUTBOUND OVERVIEW 
    
    c_out = LineChart()
    c_out.title = "Outbound Calls (Celebree)"
    c_out.style = 4
    c_out.y_axis.title = 'CALLS'
    c_out.x_axis.title = 'TIMELINE'
    
    
    c_out.height = 21
    c_out.width = 29
    
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
    data = Reference(ws, min_col=1, min_row=11, max_col=len(table_months), max_row=12)   
    c_out.add_data(data, titles_from_data=True, from_rows=True)
    c_out.set_categories(labels)
                
    
    # VISITS SCHEDULED OVERVIEW
    
    c_vis = LineChart()
    c_vis.title = "Visits Scheduled (Celebree)"
    c_vis.style = 5
    c_vis.y_axis.title = 'CALLS'
    c_vis.x_axis.title = 'TIMELINE'
    
    c_vis.height = 21
    c_vis.width = 29
    
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
    data = Reference(ws, min_col=1, min_row=14, max_col=len(table_months), max_row=15)   
    c_vis.add_data(data, titles_from_data=True, from_rows=True)
    c_vis.set_categories(labels)
                
    
    # ENROLLED OVERVIEW
    
    c_enr = LineChart()
    c_enr.title = "Enrolled (Celebree)"
    c_enr.style = 6
    c_enr.y_axis.title = 'CALLS'
    c_enr.x_axis.title = 'TIMELINE'
    
    c_enr.height = 21
    c_enr.width = 29
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
    data = Reference(ws, min_col=1, min_row=17, max_col=len(table_months), max_row=18)   
    c_enr.add_data(data, titles_from_data=True, from_rows=True)
    c_enr.set_categories(labels)
                
    
    # FTE OVERVIEW
    
    c_fte = LineChart()
    c_fte.title = "FTE Enrolled (Celebree)"
    c_fte.style = 7
    c_fte.y_axis.title = 'CALLS'
    c_fte.x_axis.title = 'TIMELINE'
    
    c_fte.height = 21
    c_fte.width = 29
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
    data = Reference(ws, min_col=1, min_row=20, max_col=len(table_months), max_row=21)   
    c_fte.add_data(data, titles_from_data=True, from_rows=True)
    c_fte.set_categories(labels)
    
    
    # GENERAL OVERVIEW ChART (Caliday)
                
    ca_all = LineChart()
    ca_all.title = "Client Conversion Pipeline Chart (Caliday)"
    ca_all.style = 2
    ca_all.y_axis.title = 'CALLS'
    ca_all.x_axis.title = 'TIMELINE'
    
    ca_all.height = 21
    ca_all.width = 29
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
    data = Reference(ws, min_col=1, min_row=32, max_col=len(table_months), max_row=35)   
    ca_all.add_data(data, titles_from_data=True, from_rows=True)
    ca_all.set_categories(labels)
    
    
    # INBOUND OVERVIEW 
    
    ca_inb = LineChart()
    ca_inb.title = "Inbound Calls (Caliday)"
    ca_inb.style = 3
    ca_inb.y_axis.title = 'CALLS'
    ca_inb.x_axis.title = 'TIMELINE'
    
    ca_inb.height = 21
    ca_inb.width = 29
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
    data = Reference(ws, min_col=1, min_row=37, max_col=len(table_months), max_row=38)   
    ca_inb.add_data(data, titles_from_data=True, from_rows=True)
    ca_inb.set_categories(labels)
    
    
    # OUTBOUND OVERVIEW 
    
    ca_out = LineChart()
    ca_out.title = "Outbound Calls (Caliday)"
    ca_out.style = 4
    ca_out.y_axis.title = 'CALLS'
    ca_out.x_axis.title = 'TIMELINE'
    
    ca_out.height = 21
    ca_out.width = 29
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
    data = Reference(ws, min_col=1, min_row=40, max_col=len(table_months), max_row=41)   
    ca_out.add_data(data, titles_from_data=True, from_rows=True)
    ca_out.set_categories(labels)
                
    
    # VISITS SCHEDULED OVERVIEW
    
    ca_vis = LineChart()
    ca_vis.title = "Links Sent (Caliday)"
    ca_vis.style = 5
    ca_vis.y_axis.title = 'CALLS'
    ca_vis.x_axis.title = 'TIMELINE'
    
    ca_vis.height = 21
    ca_vis.width = 29
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
    data = Reference(ws, min_col=1, min_row=43, max_col=len(table_months), max_row=44)   
    ca_vis.add_data(data, titles_from_data=True, from_rows=True)
    ca_vis.set_categories(labels)
                
    
    # ENROLLED OVERVIEW
    
    ca_enr = LineChart()
    ca_enr.title = "Enrolled (Caliday)"
    ca_enr.style = 6
    ca_enr.y_axis.title = 'CALLS'
    ca_enr.x_axis.title = 'TIMELINE'
    
    ca_enr.height = 21
    ca_enr.width = 29
    
    labels = Reference(ws, min_row=1, max_row=1, min_col=2, max_col=len(table_months)) 
    data = Reference(ws, min_col=1, min_row=46, max_col=len(table_months), max_row=47)   
    ca_enr.add_data(data, titles_from_data=True, from_rows=True)
    ca_enr.set_categories(labels)
    
    
    ws = wb["OVERVIEW"]
    ws.sheet_view.zoomScale = 80
    
    
    ws.add_chart(c_all, "a1")
    ws.add_chart(c_inb, "a41")
    ws.add_chart(c_out, "A81")
    ws.add_chart(c_vis, "a121")
    ws.add_chart(c_enr, "A161")
    ws.add_chart(c_fte, "a201")
    
    ws.add_chart(ca_all, "s1")
    ws.add_chart(ca_inb, "s41")
    ws.add_chart(ca_out, "s81")
    ws.add_chart(ca_vis, "s121")
    ws.add_chart(ca_enr, "s161")


       
    wb.save("output1.xlsx")

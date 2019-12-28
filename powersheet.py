# use python-tabulate to create tables for the powersheet.
# use train symbols as a list with attached engines inside the list
# Grab txt files for unit characteristics to make printing easier
from datetime import datetime
from datetime import date
from datetime import timedelta
from openpyxl import load_workbook
from openpyxl import Workbook
import requests
from lxml import html
from bs4 import BeautifulSoup
import getpass
import os
import sys
#import re

today = date.today()
curDate = today.strftime("%m-%d-%Y")
fromList = ' '
#wb = load_workbook(filename = 'Enola_Powersheet_'+curDate+'.xlsx')
wb = load_workbook(filename = 'Enola Powersheet 12-14-2019 1430.xlsx')
#current_sheet = wb[curDate]
current_sheet = wb['12-14-2019']

#print('Currently working with the',curDate,'worksheet...')
print('Using a static sheet currently')
print('At the moment this is only being built for C-trick\n\n')
print('Add a turnover function to automatically fill that out')


print("************-| LD_50 Locomotive Powersheet Mutilator |-**************")
def menu():

    choice = input("""
                
                A: Inbound and Outbound Report
                B: Outbound Trains
                C: Change Powersheet
                D: Save Powersheet
                E: Rob Peter
                F: Pay Paul
                G: Search for open engines
                H: Identify shoppers
                I: Make Work Packets
                J: Scrape Unit Information


                Q: Quit/Log Out

                Please enter your choice: """)

    if choice == "A" or choice == "a":
        dispatchReport()
    elif choice == "B" or choice == "b":
        fromBuilt()
    elif choice == "C" or choice == "c":
        appendBuild()
    elif choice == "D" or choice == "d":
        savePowersheet()
    elif choice == "E" or choice == "e":
        robPeter()
    elif choice == "F" or choice == "f":
        payPaul()
    elif choice == 'G' or choice == 'g':
        openEngines()
    elif choice == 'H' or choice == 'h':
        shoppers()
    elif choice == 'I' or choice == 'i':
        create_packets()
    elif choice == 'J' or choice == 'j':
        scrape()
    elif choice == "Q" or choice == "q":
        sys.exit
    else:
        print("You must only select either A,B,C,D,E,F,G or Q.")
        print("Please try again\n")
        menu()



def create_packets():
    USERNAME = input('\nLMIS Username: ')
    PASSWORD = getpass.getpass('LMIS Password: ')

    print('LD_50 Scrape')
    print('Author: Sean Robinson, SGL, Enola Diesel')
    print('Welcome to the LMIS Scraper...\n')

    locomotive_numbers = input('Enter locomotive numbers separated by a space: ')
    locomotive_list = locomotive_numbers.split(" ")
    date = today.strftime("%m-%d-%Y")
    unitInfo = []
    scheduled = []
    scheduled_date = []
    scheduled_tasks = []
    scheduled_task_dates = []
    # login page for LMIS
    LOGIN_URL = "https://www2.nscorp.com/mech0000/login.lmis"

    payload = {
        "username": USERNAME, 
        "pass1": PASSWORD, 
    }

    # keeps us logged into the session
    session_requests = requests.session()
    result = session_requests.get(LOGIN_URL)

    # Login
    result = session_requests.post(LOGIN_URL, data = payload, headers = dict(referer = LOGIN_URL))
     
    # Loop over the input list and scrape the work orders
    for x in locomotive_list:
        #locomotive_characteristics = ('https://www2.nscorp.com/mech0000/Locomotive?frame=frm&init=NS&nbr=9355&callingScreen=SHOPGRID')
        #unit_information_report = ("https://www2.nscorp.com/mech0000/displayReports.lmis?nme=http://mechanical.nscorp.com/loc_Info/reports/Unit_Info_Reports/NS000000"+x+".txt")
        #scheduled_maintenance_dates = ('https://www2.nscorp.com/mech0000/SmDueDates.lmis?action=S&callingScreen=OUTWRKOR&unitinit=NS&unitnumber=0000009355&inclsmi=N')
        #unit_in_shop_by_reason = ('https://www2.nscorp.com/mech0000/unitshopreason.lmis')
        #shop_it_units = ('https://www2.nscorp.com/mech0000/unitshopreasonitdetail.lmis?Shop=ENO&Reason=')
        LMIS_URL = ("https://www2.nscorp.com/mech0000/OutstandingWorkOrders.lmis?pageprocess=VT&locoinit=NS&loconbr=000000"+x+"&notfromshp=N&readonly=N&shop=%20%20%20&attachonly=N&updateact=N&searchbox=Y&reqFromModule=")
        scheduled_dates_url ="https://www2.nscorp.com/mech0000/SmDueDates.lmis?action=S&callingScreen=OUTWRKOR&unitinit=NS&unitnumber=000000"+x+"&inclsmi=N"
        
        result = session_requests.get(LMIS_URL, headers = dict(referer = LMIS_URL))
        scheduled_result = session_requests.get(scheduled_dates_url, headers = dict(referer = scheduled_dates_url))

        #unit_result = session_requests.get(unit_information_report, headers = dict(referer = unit_information_report))
        soup = BeautifulSoup(result.content, 'lxml')
        scheduled_soup = BeautifulSoup(scheduled_result.content, 'lxml')
        
        

        #unit_soup = BeautifulSoup(unit_result.content, 'lxml')

        # add PTC health when able
        # add DP also
        
        # scheduled due dates 
        # items are hidden on the due dates page
        # they are not static and will have to be searched through
        #print('test fire!--------->', scheduled_soup.find("input", {"name":"hTaskType0"})['value'])
        # scanning due dates

        #this print the category and due date, make this into a list or table?
        for task in range(0, 19):
            due_task = scheduled_soup.find("input", {"name":"hTaskType"+str(task)})['value']
            due_date_str = scheduled_soup.find("input", {"name":"hNextDueDate"+str(task)})['value']
            due_date = datetime.strptime(due_date_str, '%m-%d-%Y')
            mi_date = (datetime.today() + timedelta(6))
            if due_date < mi_date:
                scheduled_tasks.append(due_task)
                scheduled_task_dates.append(due_date_str)
            #compare dates and add due dates to a list
            #check list for comparison to see if something exists (lube, labs, af, etc)
            #if it exists, make a variable so it can be added to the work packet
            #scheduled_task_item = due_task,due_date_str     
            #scheduled_tasks.append(scheduled_task_item)
        table_format(scheduled_tasks, scheduled_task_dates, 'Tasks Due', 'Due Date')

        #scheduled.append(list(zip(scheduled_tasks,scheduled_task_dates)))
        scheduled.append(list(scheduled_tasks))
        scheduled_date.append(list(scheduled_task_dates))
        scheduled_tasks.clear()
        scheduled_task_dates.clear()
        #print(scheduled)
        #print(scheduled_date)
        #print(str(scheduled))
        # TS,LS,LB,1Y,2Y,N6,M5,M6,AF,EV,MR,CS,HB,AN,AB,RS, N2,N4,N5,M7,M2,M3,M5,RS(tape)
        #maintenance_dates(scheduled_tasks, scheduled_task_dates)

        print('FIXME: add ptc health to end of PTC line...add DP to report\n\n')
        print('----',x,'----')
        if (soup.find("input", {"name":"hModel"})) is not None:
            model = soup.find("input", {"name":"hModel"})['value']
            print('Model: ' + model)
        if (soup.find("input", {"name":"hPtc"})) is not None:
            ptc = soup.find("input", {"name":"hPtc"})['value']
            print('PTC: ' + ptc)
        if (soup.find("input", {"name":"hEM"})) is not None:
            em = soup.find("input", {"name":"hEM"})['value']
            print('EM: ' + em)
        if (soup.find("input", {"name":"hCs"})) is not None:
            cabs = soup.find("input", {"name":"hCs"})['value']
            print('CS: ' + cabs)
        if (soup.find("input", {"name":"hLSL"})) is not None:
            lsl = soup.find("input", {"name":"hLSL"})['value']
            print('LSL: ' + lsl)
        if (soup.find("input", {"name":"hRelIu"})) is not None:
            relInd = soup.find("input", {"name":"hRelIu"})['value']
            print('Reliability: ' + relInd)
        if (soup.find("input", {"name":"hEquivAxl"})) is not None:
            group = soup.find("input", {"name":"hEquivAxl"})['value']
            print('Power Group: ' + group)
        if (soup.find("input", {"name":"hPropDue"})) is not None:
            fra = soup.find("input", {"name":"hPropDue"})['value']
            print('FRA Due: ' + fra)
        if (soup.find("input", {"name":"hEpaDead"})) is not None:
            epa = soup.find("input", {"name":"hEpaDead"})['value']
            print('EPA Due: ' + epa)
        if (soup.find("input", {"name":"hLubeDue"})) is not None:
            lube = soup.find("input", {"name":"hLubeDue"})['value']
            print('Lube Due: ' + lube)
        if (soup.find("input", {"name":"hCabS"})) is not None:
            csDue = soup.find("input", {"name":"hCabS"})['value']
            print('Cab Signals Due: ' + csDue)
        else:
            csDue = '-'
        if (soup.find("input", {"name":"hFc"})) is not None:
            fuelCap = soup.find("input", {"name":"hFc"})['value']
            print('Fuel Capacity: ' + fuelCap.lstrip("0"))
        if (soup.find("input", {"name:":"hNextFraAirFlowMeter"})) is not None:
            airFlow = soup.find("input", {"name:":"hNextFraAirFlowMeter"})['value'] 
        else:
            airFlow = '-'
        
        locomotive_Info = x,date,fra,epa,'Y',csDue
        unitInfo.append(locomotive_Info)



    correctInfo=input("\nIs the information correct? (y/n) ")
    if correctInfo == 'y' or correctInfo == 'Y':
        mi_starting_cell = 24
        ur_starting_cell = 23
        miCover = load_workbook(filename="MIPacketCover.xlsx")
        urCover = load_workbook(filename="URPacketCover.xlsx")
        j = 0
        for info in unitInfo:
            maint = input('Is '+info[0]+' a maintenance unit? (y/n) ')
            if maint == "y" or maint == "Y":
                print('Saving cover for Unit #: '+info[0]+'.')
                packet = miCover.copy_worksheet(miCover["MI Cover Sheet"])
                packet.title=info[0]
                # trying to the new function
                # maintenance_dates(scheduled_tasks, scheduled_task_dates, packet)
                maintenance_dates(scheduled[int(j)], scheduled_date[int(j)],
                                 packet)
                worksheet_tasks(packet, mi_starting_cell)
                packet.cell(row=2, column=1).value = info[0]
                packet.cell(row=1, column=6).value = info[1]
                packet.cell(row=5, column=3).value = info[2]
                #packet.cell(row=7, column=3).value = info[3]
                #packet.cell(row=3, column=6).value = 'Y'
                packet.cell(row=4, column=6).value = 'Y'
                #packet.cell(row=5, column=6).value = 'Y'
                #packet.cell(row=6, column=6).value = info[5]
                #packet.cell(row=7, column=6).value = airFlow
                #print(packet.cell(row=2, column=1).value)
                j += 1
            elif maint == 'n' or maint == 'N':
                print('Saving cover for Unit #: '+info[0]+'.')
                packet = urCover.copy_worksheet(urCover["UR Cover Sheet"])
                packet.title=info[0]
                packet.cell(row=2, column=1).value = info[0]
                packet.cell(row=1, column=6).value = info[1]
                packet.cell(row=5, column=3).value = info[2]
                maintenance_dates(scheduled[int(j)], scheduled_date[int(j)],
                                  packet)
                #maintenance_dates(scheduled_tasks, scheduled_task_dates, packet)
                worksheet_tasks(packet, ur_starting_cell)
                packet.cell(row=4, column=6).value = 'Y'
                j += 1
                #print(packet.cell(row=2, column=1).value)               
        #del urCover['UR Cover Sheet']
        #urCover.save('UR_Units_CoverSheets_'+curDate+'.xlsx')
        urCover.save('UR_CoverSheets.xlsx')
        #del miCover['MI Cover Sheet']
        miCover.save('MI_CoverSheets.xlsx')
        # miCover.save('MI_Unit_CoverSheets_'+curDate+'.xlsx')
        #mr = soup(text=re.compile('MR')))
      
    menu()

def scrape():
    USERNAME = input('\nLMIS Username: ')
    PASSWORD = getpass.getpass('LMIS Password: ')

    print('LD_50 Scrape')
    print('Author: Sean Robinson, SGL, Enola Diesel')
    print('Welcome to the LMIS Scraper...\n')

    UNIT_NUMBERS = input('Enter locomotive numbers separated by a space: ')
    UNIT_LIST = UNIT_NUMBERS.split(" ")

    # login page for LMIS
    LOGIN_URL = "https://www2.nscorp.com/mech0000/login.lmis"

    payload = {
        "username": USERNAME, 
        "pass1": PASSWORD, 
    }

    #URL = "https://www2.nscorp.com/mech0000/OutstandingWorkOrders.lmis?pageprocess=VT&locoinit=NS&loconbr=0000009952&notfromshp=N&readonly=N&shop=%20%20%20&attachonly=N&updateact=N&searchbox=Y&reqFromModule="    
    
    # keeps us logged into the session
    session_requests = requests.session()
    result = session_requests.get(LOGIN_URL)

    # Login
    result = session_requests.post(LOGIN_URL, data = payload, headers = dict(referer = LOGIN_URL))
     
    # Loop over the input list and scrape the work orders
    for x in UNIT_LIST:
    
        LMIS_URL = "https://www2.nscorp.com/mech0000/OutstandingWorkOrders.lmis?pageprocess=VT&locoinit=NS&loconbr=000000"+x+"&notfromshp=N&readonly=N&shop=%20%20%20&attachonly=N&updateact=N&searchbox=Y&reqFromModule="
        result = session_requests.get(LMIS_URL, headers = dict(referer = LMIS_URL))
        #with open("/home/robbie/Code/Work/"+x+".html", 'wb') as file:
        #    file.write(result.content)

        soup = BeautifulSoup(result.content, 'lxml')

        # add PTC health when able
        # add DP also

        print('FIXME: add ptc health to end of PTC line...add DP to report\n\n')
        print('----',x,'----')
        if (soup.find("input", {"name":"hModel"})) is not None:
            model = soup.find("input", {"name":"hModel"})['value']
            print('Model: ' + model)
        if (soup.find("input", {"name":"hPtc"})) is not None:
            ptc = soup.find("input", {"name":"hPtc"})['value']
            print('PTC: ' + ptc)
        if (soup.find("input", {"name":"hEM"})) is not None:
            em = soup.find("input", {"name":"hEM"})['value']
            print('EM: ' + em)
        if (soup.find("input", {"name":"hCs"})) is not None:
            cabs = soup.find("input", {"name":"hCs"})['value']
            print('CS: ' + cabs)
        if (soup.find("input", {"name":"hLSL"})) is not None:
            lsl = soup.find("input", {"name":"hLSL"})['value']
            print('LSL: ' + lsl)
        if (soup.find("input", {"name":"hRelIu"})) is not None:
            relInd = soup.find("input", {"name":"hRelIu"})['value']
            print('Reliability: ' + relInd)
        if (soup.find("input", {"name":"hEquivAxl"})) is not None:
            group = soup.find("input", {"name":"hEquivAxl"})['value']
            print('Power Group: ' + group)
        if (soup.find("input", {"name":"hPropDue"})) is not None:
            fra = soup.find("input", {"name":"hPropDue"})['value']
            print('FRA Due: ' + fra)
        if (soup.find("input", {"name":"hEpaDead"})) is not None:
            epa = soup.find("input", {"name":"hEpaDead"})['value']
            print('EPA Due: ' + epa)
        if (soup.find("input", {"name":"hLubeDue"})) is not None:
            lube = soup.find("input", {"name":"hLubeDue"})['value']
            print('Lube Due: ' + lube)
        if (soup.find("input", {"name":"hCabS"})) is not None:
            csDue = soup.find("input", {"name":"hCabS"})['value']
            print('Cab Signals Due: ' + csDue)
        if (soup.find("input", {"name":"hFc"})) is not None:
            fuelCap = soup.find("input", {"name":"hFc"})['value']
            print('Fuel Capacity: ' + fuelCap.lstrip("0"))

    menu()

def worksheet_tasks(packet, cell):
    #ur=32 38 33 39
    work_list = []
    work_header = []
    while True:
        header = input('Input worksheet header: ')
        if header == 'Q' or header =='q' or header == '':
            break
        work_header.append(header)
        print(work_header)
        work_task = input('Input worksheet task: ')
        if work_task == 'Q' or work_task == 'q' or work_task == '':
            break
        work_list.append(work_task)
        print(work_list)

    for head in work_header:
        work_cell_iterator = int(cell)
        while packet.cell(row=work_cell_iterator, column=2).value is not None:
            if work_cell_iterator == 32:
                work_cell_iterator = int(work_cell_iterator) + 6
            elif work_cell_iterator == 33:
                work_cell_iterator = int(work_cell_iterator) + 5
            else:
                work_cell_iterator = int(work_cell_iterator) + 3
        print('Adding header: '+head)
        packet.cell(row=work_cell_iterator, column=2).value = head

    for task in work_list:
        work_cell_iterator = int(cell) + 1
        while packet.cell(row=work_cell_iterator, column=2).value is not None:
            #print(work_cell_iterator)
            if work_cell_iterator == 33:
                work_cell_iterator = int(work_cell_iterator) + 6
            elif work_cell_iterator == 34:
                work_cell_iterator = int(work_cell_iterator) + 5
            else:
                work_cell_iterator = int(work_cell_iterator) + 3
        print('Adding task: '+task)
        packet.cell(row=work_cell_iterator, column=2).value = task
def maintenance_dates(tasks, due_dates, packet):
    mi_due_cell = packet.cell(row=6, column=3).value
    epa_due_cell = packet.cell(row=7, column=3).value
    epa_due = ''
    mi_due = ''
    if 'LS' in tasks:
        #print('Samples due.')
        task_index = tasks.index('LS')
        packet.cell(row=3, column=6).value = 'Y'
    if 'AF' in tasks:
        #print('Airflow due.')
        task_index = tasks.index('AF')
        #air_flow = due_dates[task_index]
        packet.cell(row=7, column=6).value = 'Y'
    if 'CS' in tasks:
        #print('Cabs due.')
        packet.cell(row=6, column=6).value = 'Y'
    if 'LB' in tasks:
        #print('Lube due.')
        packet.cell(row=5, column=6).value = 'Y'
    if '1Y' in tasks and '2Y' in tasks:
        #print('EPA: 1Y, 2Y')
        packet.cell(row=7, column=3).value = epa_dua+'1Y, 2Y'
    elif '1Y' in tasks:
        #print('1Y EPA')
        packet.cell(row=7, column=3).value = epa_due+'1Y'
    elif '2Y' in tasks:
        #print('EPA: 2Y')
        packet.cell(row=7, column=3).value = epa_due+'2Y'
    if 'M5' in tasks:
        #print('EPA: M5')
        packet.cell(row=7, column=3).value = epa_due+'M5'
    if 'M6' in tasks:
        packet.cell(row=7, column=3).value = epa_due+'M6'
        #print('EPA: M6')
    if 'M7' in tasks:
        packet.cell(row=7, column=3).value = epa_due+'M7'
        #print('EPA: M7')
    if 'N6' in tasks:
        packet.cell(row=7, column=3).value = epa_due+',N6'
    if 'MR' in tasks:
        packet.cell(row=6, column=3).value = '6mo'
    if 'AN' in tasks:
        #print('12mo.')
        packet.cell(row=6, column=3).value = '12mo'
    if 'AB' in tasks:
        #print('Air change.')
        packet.cell(row=6, column=3).value = packet.cell(row=6,
                                                         column=3).value+', Air'
def writeMultColumns(row, column1, column2, robbed, paid):
    r = row
    col1 = column1
    col2 = column2
    t = robbed
    p = paid
    #u = units

    print('|---------- Writing Data ----------|')

    #check for empty cells
    emptyCol1 = current_sheet.cell(row=r, column=col1).value
    emptyCol2 = current_sheet.cell(row=r, column=col2).value 

    if emptyCol1 and emptyCol2 is None:
        emptyCol1 = t
        emptyCol2 = p

def readSingleColumn(rowStart, rowEnd, col):
    row_start = rowStart
    row_end= rowEnd
    column = col
    this_list= []
    for row in range(row_start, row_end):
        cell_value = current_sheet.cell(row=row, column=column).value
        if cell_value is not None:
            this_list.append(cell_value)
        elif cell_value is None:
            this_list.append('-')
    return this_list

def readMultColumns(rowStart, rowEnd, colStart, colEnd):
    rs = rowStart
    re = rowEnd
    cs = colStart
    ce = colEnd
    for i in range(rs, re):
        x = current_sheet.cell(row=i, column=cs).value
        z = current_sheet.cell(row=i, column=ce).value
        if x is None and z is not None:
            print(x,z)
        elif x is int and x is not None:
            print(x,'-->',z)
        elif x is not None:
            print(str(x)[:3],'-->',z)
            #print(x,'-->',z)

def readMultColumnsTable(rowStart, rowEnd, colOne, colTwo):
    row_start = rowStart
    row_end = rowEnd
    column_one = colOne
    column_two = colTwo
    list_one = []
    list_two = []

    for cell in range(row_start, row_end):
        column1 = current_sheet.cell(row=cell, column=column_one).value
        column2 = current_sheet.cell(row=cell, column=column_two).value
        if column1 is None and column2 is not None:
            list_one.append('-')
            list_two.append(column2)
        elif column1 is int and column2 is not None:
            list_one.append(column1)
            list_two.append(column2)
        elif column1 is not None and column2 is None:
            list_one.append(str(column1)[:3]+' ---------->')
            list_two.append('-')
        elif column1 is not None:
            list_one.append(str(column1)[:3]+' ---------->')
            list_two.append(column2)
    return list_one, list_two

def table_format(list_one, list_two, label_one, label_two):
    fmt = '{:<8}{:<20}{}'
    print(fmt.format('', label_one, label_two))
    for stuff, (row,column) in enumerate(zip(list_one, list_two)):
        print(fmt.format((stuff + 1), row, column))
        
def table_format_three(list_one, list_two, list_three, label_one, label_two, label_three):
    fmt = '{:<8}{:<20}{:<50}{}'
    print(fmt.format('', label_one, label_two, label_three))
    for stuff, (list_a, list_b, list_c) in enumerate(zip(list_one, list_two, list_three)):
        print(fmt.format((stuff + 1), list_a, list_b, list_c))

def openEngines():
    print('\nSearching for Open Locomotives on Current Sheet.....')
    
    inboundUnits=[]
    for rows in range (4, 27):
        row = current_sheet.cell(row=rows, column = 3).value
        if row is not None:
            rowSlicer = row.split()
            slash = '/'
            while slash in rowSlicer: rowSlicer.remove(slash)
            newList = [s[:4] for s in rowSlicer if s[:3].isdigit()]
            inboundUnits.extend(newList)
    usedUnits=[]
    for rows in range(4, 27):
        row = current_sheet.cell(row=rows, column=14).value
        if row is not None:
            rowSlicer = row.split()
            slash = '/'
            while slash in rowSlicer: rowSlicer.remove(slash)
            newList = [s[:4] for s in rowSlicer if s[:3].isdigit()]
            usedUnits.extend(newList)

    openUnits = [s for s in inboundUnits if s not in usedUnits]
    print('Search complete...\n\nOpen Locomtovies:',openUnits,'\n')
    menu()

def robPeter():
    print('\n|--------------- Robbing Peter ---------------|\n')
    
    stolenTrain = input('What train are we stealing from?  ')
    stolenUnits = input('What units are being stolen?  ')

    stolenList = stolenUnits.split()
    print(stolenList)
    
    payTrain = input('What train are we paying '+stolenTrain+' to? (ex. 15T - OUTBOUND)  ')

    #check inbound side for an empty slot to fill in the robbed train
    #then insert the information 
    for train in range(15, 28):
        checkTrain = current_sheet.cell(row=train, column=1).value
        checkUnits = current_sheet.cell(row=train, column=3).value
        if checkTrain is None and checkUnits is None:
            if stolenTrain.find('.') >= 0:
                datedTrain = stolenTrain.split('.')
                current_sheet.cell(row=train,column=1).value=datedTrain[0]
                current_sheet.cell(row=train, column=2).value=datedTrain[1]
                current_sheet.cell(row=train, column=3).value = stolenUnits
            else:
                current_sheet.cell(row=train, column=1).value = stolenTrain
                current_sheet.cell(row=train, column=3).value = stolenUnits
            current_sheet.cell(row=train, column=8).value = payTrain
            print(checkTrain,':',checkUnits)
            break

    #print('How many units will be needed to replace robbed power?')
    needed=input('How many units will be needed to replace robbed power? (NEED X)  ')
    
    #check for empty rows in the outbound extras 
    for train in range(15, 28):
        checkTrain = current_sheet.cell(row=train, column=12).value
        checkUnits = current_sheet.cell(row=train, column=14).value
        if checkTrain is None and checkUnits is None:
            current_sheet.cell(row=train, column=12).value = stolenTrain
            current_sheet.cell(row=train, column=14).value = needed
            print(checkTrain,':',checkUnits)
            break
    dispatchReport()

def payPaul():
    print('\n|--------------- Paying Paul ---------------|\n')
    
    payTrain = input('What train are we paying back?  ')
    payUnits = input('What units are being given?  ')

    payList = payUnits.split()
    print(payUnits)
    
    fromTrain = input('What train are we paying '+payTrain+' from? ')

    for train in range(16, 28):
        checkTrain = current_sheet.cell(row=train, column=12).value
        checkUnits = current_sheet.cell(row=train, column=14).value
        if checkTrain == payTrain:
            current_sheet.cell(row=train, column=14).value = payUnits
            current_sheet.cell(row=train, column=19).value = fromTrain
            print(checkTrain,':',checkUnits)
            break
    #print('Repaid check')

    dispatchReport()

def shoppers():
    shoppers = input('\nEnter shopped units coming in: ')
    print(shoppers)
    print('FIXME: Use this function to color shoppers red and add them to the assignemnts sheets')
    menu()

def searchPower(newPower, fromRow):
    power = [x.strip() for x in newPower.split('/')]
    row=fromRow
    units=[]
    for i in power:
        x = i[:4]
        for j in range (4, 27):
            # assigning y to train symbols
            y = current_sheet.cell(row=j, column=1).value
            if y is not None:
                a = str(y)[:3]
                z = current_sheet.cell(row=j, column=3).value
                if z is not None:
                    zlist = z.strip()
                    search = zlist.find(x)
                    if search >= 0:
                        units.append(a)
    remove_adjacent(units)
    fromList = (' / '.join(units))
    print('From: ',fromList)
    current_sheet.cell(row=row, column=21).value = fromList

def remove_adjacent(seq): # works on any sequence, not just on numbers
    i = 1
    n = len(seq)
    while i < n: # avoid calling len(seq) each time around
        if seq[i] == seq[i-1]:
            del seq[i]
            # value returned by seq.pop(i) is ignored; slower than del seq[i]
            n -= 1
        else:
            i += 1

def appendBuild():
    print('\n|*************** Changing Power ***************|\n')

    print('\n----------| Inbound Trains |------------\n')
    inbound_train,inbound_power = readMultColumnsTable(4, 16, 1, 3)
    table_format(inbound_train, inbound_power, 'Train Symbol', 'Power')

    print('\n----------| Outbound Trains |------------\n')
    outbound_train, outbound_power = readMultColumnsTable(4, 16, 12, 14)
    table_format(outbound_train, outbound_power, 'Train Symbol', 'Power')

    selectPower = input('\n\nWhat train would you like to change?  ')
    # row will be the current row, we have to add 3 to get down to the correct cell
 
    row = int(selectPower) + int(3)
    old_build = current_sheet.cell(row=int(row), column=14).value
    
    print('\n|---------- Appending build for',current_sheet.cell(row=row, column=12).value,'----------|\n')
    print('Current build: ', old_build)
    
    newBuild = input('New build:  ')
    confirmBuild = input('You would like to change the build to: ' + newBuild + ' ? (y/n)  ')

    if confirmBuild == 'y' or confirmBuild =="Y":
        print('\nChanging power from ', old_build, ' to :', newBuild,' .')
        current_sheet.cell(row=row, column=14).value = newBuild
        print('Power changed to', current_sheet.cell(row=row, column=14).value, ' .\n')
        
        #Editing the "From" category on Outbound Trains
        newBuildValue = current_sheet.cell(row=row, column=14).value
        print('Consist:',newBuildValue)
        
        #Searches the inbound columns to find where power came from
        searchPower(newBuildValue, row)
        print(current_sheet.cell(row=row, column=12).value,' changed to: ',newBuild, ' FROM:', current_sheet.cell(row=row, column=21).value, '.\n')
        
    else:
        print('No changes made.\n')
    menu()

def dispatchReport():
    print('\n---------- Inbound Trains ----------\n')
    inbound_train,inbound_power = (readMultColumnsTable(4, 15, 1, 3))
    table_format(inbound_train, inbound_power, 'Train Symbol','Power')
    print('\n---------- Outbound Trains ----------\n')
    outbound_trains, outbound_power = (readMultColumnsTable(4, 16, 12, 14))
    table_format(outbound_trains, outbound_power, 'Train Symbol','Power')
    print('\n---------- Extra Inbounds ----------\n')
    extra_inbounds, extra_inbound_power = (readMultColumnsTable(15, 28, 1, 3))
    table_format(extra_inbounds, extra_inbound_power, 'Train Symbol','Power')
    print('\n---------- Extra Outbounds ----------\n')
    extra_outbound, extra_outbound_power = (readMultColumnsTable(16, 28, 12, 14))
    table_format(extra_outbound, extra_outbound_power, 'Train Symbol','Power')
    print('\n')
    menu()
    
def fromBuilt():

    outbound_trains, outbound_power = (readMultColumnsTable(4, 18, 12, 14))
    from_symbol = readSingleColumn(4, 18, 21)
    #table_format(outbound_trains, outbound_power, 'Train Symbol', 'Power')
    table_format_three(outbound_trains, outbound_power, from_symbol, 'Train Symbol', 'Power', 'From')
    print('\n')
    #for i in range(4, 15):
    #    x = current_sheet.cell(row=i, column=12).value
    #    y = x[:3]
    #    z = current_sheet.cell(row=i, column=14).value
    #    a = current_sheet.cell(row=i, column=21).value
    #    print(y,': ',z)
    #    print('From:',a,'\n')
    #    #print('----: ',a,'\n')
    #print('\n')
    menu()

def savePowersheet():
    print('Saving workboot.....')
    #wb.save('/home/'+getpass.getuser()+'/Documents/test.xlsx')
    wb.template = False
    wb.save('test.xlsx')
    menu()

if __name__ == '__main__':
    menu()

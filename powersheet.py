from datetime import date
from openpyxl import load_workbook
import requests
from lxml import html
from bs4 import BeautifulSoup
import getpass
import os
import sys

#os.chdir('/home/robbie/Code/Work/excel/')
#wb = load_workbook(filename = 'Powersheet_Template2.xlsx')

today = date.today()
curDate = today.strftime("%m-%d-%Y")

#curDate = '12-02-2019'

fromList = ' '
wb = load_workbook(filename = 'Enola_Powersheet_'+curDate+'.xlsx')
curSheet = wb[curDate]

print('Currently working with the',curDate,'worksheet...')
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
        createPackets()
    elif choice == 'J' or choice == 'j':
        scrape()
    elif choice == "Q" or choice == "q":
        sys.exit
    else:
        print("You must only select either A,B,C,D,E,F,G or Q.")
        print("Please try again\n")
        menu()



def createPackets():
    USERNAME = input('\nLMIS Username: ')
    PASSWORD = getpass.getpass('LMIS Password: ')

    print('LD_50 Scrape')
    print('Author: Sean Robinson, SGL, Enola Diesel')
    print('Welcome to the LMIS Scraper...\n')

    UNIT_NUMBERS = input('Enter locomotive numbers separated by a space: ')
    UNIT_LIST = UNIT_NUMBERS.split(" ")
    unitInfo = []
    date = today.strftime("%m-%d")
    # login page for LMIS
    LOGIN_URL = "https://www2.nscorp.com/mech0000/login.lmis"

    payload = {
        "username": USERNAME, 
        "pass1": PASSWORD, 
    }

    URL = "https://www2.nscorp.com/mech0000/OutstandingWorkOrders.lmis?pageprocess=VT&locoinit=NS&loconbr=0000009952&notfromshp=N&readonly=N&shop=%20%20%20&attachonly=N&updateact=N&searchbox=Y&reqFromModule="    
    
    # keeps us logged into the session
    session_requests = requests.session()
    result = session_requests.get(LOGIN_URL)

    # Login
    result = session_requests.post(LOGIN_URL, data = payload, headers = dict(referer = LOGIN_URL))
     
    # Loop over the input list and scrape the work orders
    for x in UNIT_LIST:
        SCRAPE_URL = "https://www2.nscorp.com/mech0000/OutstandingWorkOrders.lmis?pageprocess=VT&locoinit=NS&loconbr=000000"+x+"&notfromshp=N&readonly=N&shop=%20%20%20&attachonly=N&updateact=N&searchbox=Y&reqFromModule="
        result = session_requests.get(SCRAPE_URL, headers = dict(referer = SCRAPE_URL))
        
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
        # Make a string that holds all this information per locomotive. 
        # Make packets as a batch after all info has been pulled 
        # (LMIS looks to drop out after stopping to ask about packets)
        #locoInfo = x,',',date,',',fra,',',epa,',','Y',',',cabs

        locoInfo = x,date,fra,epa,'Y',csDue
        unitInfo.append(locoInfo)
    correctInfo=input("\nIs the information correct? (y/n) ")
    if correctInfo == 'y' or correctInfo == 'Y':
        for info in unitInfo:
            print('Saving cover for unit.')
            URpacket = load_workbook(filename='URPacketCover.xlsx')
            urpacket = URpacket.active        
            urpacket.cell(row=2, column=1).value = info[0]
            urpacket.cell(row=1, column=6).value = info[1]
            urpacket.cell(row=5, column=3).value = info[2]
            urpacket.cell(row=7, column=3).value = info[3]
            urpacket.cell(row=4, column=6).value = info[4]
            urpacket.cell(row=6, column=6).value = info[5]
            urpacket.template = False
            URpacket.save(info[0]+'.xlsx')
    

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

    URL = "https://www2.nscorp.com/mech0000/OutstandingWorkOrders.lmis?pageprocess=VT&locoinit=NS&loconbr=0000009952&notfromshp=N&readonly=N&shop=%20%20%20&attachonly=N&updateact=N&searchbox=Y&reqFromModule="    
    
    # keeps us logged into the session
    session_requests = requests.session()
    result = session_requests.get(LOGIN_URL)

    # Login
    result = session_requests.post(LOGIN_URL, data = payload, headers = dict(referer = LOGIN_URL))
     
    # Loop over the input list and scrape the work orders
    for x in UNIT_LIST:
    
        SCRAPE_URL = "https://www2.nscorp.com/mech0000/OutstandingWorkOrders.lmis?pageprocess=VT&locoinit=NS&loconbr=000000"+x+"&notfromshp=N&readonly=N&shop=%20%20%20&attachonly=N&updateact=N&searchbox=Y&reqFromModule="
        result = session_requests.get(SCRAPE_URL, headers = dict(referer = SCRAPE_URL))
        
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

def writeMultColumns(row, column1, column2, robbed, paid):
    
    r = row
    col1 = column1
    col2 = column2
    t = robbed
    p = paid
    #u = units

    print('|---------- Writing Data ----------|')

    #check for empty cells
    emptyCol1 = curSheet.cell(row=r, column=col1).value
    emptyCol2 = curSheet.cell(row=r, column=col2).value 

    if emptyCol1 and emptyCol2 is None:
        emptyCol1 = t
        emptyCol2 = p

def readSingleColumn(rowStart, rowEnd, col):
    x = rowStart
    y = rowEnd
    c = col
    for i in range(x, y):
        z = curSheet.cell(row=i, column=c).value
        print(z)

def readMultColumns(rowStart, rowEnd, colStart, colEnd):
    rs = rowStart
    re = rowEnd
    cs = colStart
    ce = colEnd
    for i in range(rs, re):
        x = curSheet.cell(row=i, column=cs).value
        z = curSheet.cell(row=i, column=ce).value
        if x is None and z is not None:
            print(x,z)
        elif x is int and x is not None:
            print(x,'-->',z)
        elif x is not None:
            print(str(x)[:3],'-->',z)
            #print(x,'-->',z)

def openEngines():
    print('\nSearching for Open Locomotives on Current Sheet.....')
    
    inboundUnits=[]
    for rows in range (4, 27):
        row = curSheet.cell(row=rows, column = 3).value
        if row is not None:
            rowSlicer = row.split()
            slash = '/'
            while slash in rowSlicer: rowSlicer.remove(slash)
            newList = [s[:4] for s in rowSlicer if s[:3].isdigit()]
            inboundUnits.extend(newList)
    usedUnits=[]
    for rows in range(4, 27):
        row = curSheet.cell(row=rows, column=14).value
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
    #print(stolenList)
    
    payTrain = input('What train are we paying '+stolenTrain+' to? (ex. 15T - OUTBOUND)  ')

    #check inbound side for an empty slot to fill in the robbed train
    #then insert the information 
    for train in range(15, 28):
        checkTrain = curSheet.cell(row=train, column=1).value
        checkUnits = curSheet.cell(row=train, column=3).value
        if checkTrain is None and checkUnits is None:
            if stolenTrain.find('.') >= 0:
                datedTrain = stolenTrain.split('.')
                curSheet.cell(row=train,column=1).value=datedTrain[0]
                curSheet.cell(row=train, column=2).value=datedTrain[1]
                curSheet.cell(row=train, column=3).value = stolenUnits
            else:
                curSheet.cell(row=train, column=1).value = stolenTrain
                curSheet.cell(row=train, column=3).value = stolenUnits
            curSheet.cell(row=train, column=8).value = payTrain
            print(checkTrain,':',checkUnits)
            break

    #print('How many units will be needed to replace robbed power?')
    needed=input('How many units will be needed to replace robbed power? (NEED X)  ')
    
    #check for empty rows in the outbound extras 
    for train in range(15, 28):
        checkTrain = curSheet.cell(row=train, column=12).value
        checkUnits = curSheet.cell(row=train, column=14).value
        if checkTrain is None and checkUnits is None:
            curSheet.cell(row=train, column=12).value = stolenTrain
            curSheet.cell(row=train, column=14).value = needed
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
        checkTrain = curSheet.cell(row=train, column=12).value
        checkUnits = curSheet.cell(row=train, column=14).value
        if checkTrain == payTrain:
            curSheet.cell(row=train, column=14).value = payUnits
            curSheet.cell(row=train, column=19).value = fromTrain
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
            y = curSheet.cell(row=j, column=1).value
            if y is not None:
                a = str(y)[:3]
                z = curSheet.cell(row=j, column=3).value
                if z is not None:
                    zlist = z.strip()
                    search = zlist.find(x)
                    if search >= 0:
                        units.append(a)
    remove_adjacent(units)
    fromList = (' / '.join(units))
    print('From: ',fromList)
    curSheet.cell(row=row, column=21).value = fromList

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

    print('\n----------Inbound trains----------')
    for i in range(4, 16):
        x = curSheet.cell(row=i, column=1).value
        #sets y = to the first 3 letters of the string
        y = x[:3]
        z = curSheet.cell(row=i, column=3).value
        print(y,': ',z)
    print('------------------------------------\n')
    
    j = 1
    for i in range(4, 16):
        x = curSheet.cell(row=i, column=12).value
        y = x[:3]
        print(j,':',y,'-',curSheet.cell(row=i, column=14).value)
        j += 1

    selectPower = input('\n\nWhat train would you like to change?  ')
    # row will be the current row, we have to add 3 to get down to the correct cell
 
    row = int(selectPower) + int(3)
    cValue = curSheet.cell(row=int(row), column=14).value
    
    print('\n|---------- Appending build for',curSheet.cell(row=row, column=12).value,'----------|\n')
    print('Current build: ', cValue)
    
    newBuild = input('New build:  ')
    confirmBuild = input('You would like to change the build to: ' + newBuild + ' ? (y/n)  ')

    if confirmBuild == 'y' or confirmBuild =="Y":
        print('\nChanging power from ', cValue, ' to :', newBuild,' .')
        curSheet.cell(row=row, column=14).value = newBuild
        print('Power changed to', curSheet.cell(row=row, column=14).value, ' .\n')
        #Editing the "From" category on Outbound Trains
        newBuildValue = curSheet.cell(row=row, column=14).value
        print('Consist:',newBuildValue)
        
        #Searches the inbound columns to find where power came from
        searchPower(newBuildValue, row)
     
        #newFrom = input('---From: ')
        #curSheet.cell(row=row, column=21).value = newFrom
        print(curSheet.cell(row=row, column=12).value,' changed to: ',newBuild, ' FROM:', curSheet.cell(row=row, column=21).value, '.\n')
        
    else:
        print('No changes made.\n')
    menu()

def dispatchReport():
    print('\n---------- Inbound Trains ----------\n')
    readMultColumns(4, 15, 1, 3)    
    print('\n---------- Outbound Trains ----------\n')
    readMultColumns(4, 16, 12, 14)
    print('\n---------- Extra Inbounds ----------\n')
    readMultColumns(15, 28, 1, 3)
    print('\n---------- Extra Outbounds ----------\n')
    readMultColumns(16, 28, 12, 14)
    print('\n')
    menu()

def fromBuilt():
    print('\n')
    for i in range(4, 15):
        x = curSheet.cell(row=i, column=12).value
        y = x[:3]
        z = curSheet.cell(row=i, column=14).value
        a = curSheet.cell(row=i, column=21).value
        print(y,': ',z, '-->',a)
        #print('----: ',a,'\n')
    print('\n')
    menu()

def savePowersheet():
    print('Saving workboot.....')
    #wb.save('/home/'+getpass.getuser()+'/Documents/test.xlsx')
    wb.template = False
    wb.save('test.xlsx')
    menu()

if __name__ == '__main__':
    menu()
# search input without /'s
# turn new build into a list (break at spaces)
# search column for each list entry
# compare list entries
# if adjacent matches, remove one
# print members, separated by slashes


from datetime import date
from openpyxl import load_workbook
import getpass
import os
import sys

#os.chdir('/home/robbie/Code/Work/excel/')
wb = load_workbook(filename = 'Powersheet_Template2.xlsx')
today = date.today()
curDate = today.strftime("%m-%d-%Y")
curDate = '12-02-2019'
curSheet = wb[curDate]
fromList = ' '

print('Currently working with the',curDate,'worksheet...')
print('At the moment this is only being built for C-trick\n\n')
print('FIXME: Change current date variable back.....')
print('FIXME: Used 12-02-2019 for testing purposes\n\n')



def menu():
    print("************-| LD_50 Locomotive Powersheet Mutilator |-**************")
    print()

    choice = input("""
                A: Inbound and Outbound Report
                B: Outbound Trains
                C: Change Powersheet
                D: Save Powersheet
                E: Rob Peter
                F: Pay Paul
                G: Make Work Packets


                Q: Quit/Log Out

                Please enter your choice: """)

    if choice == "A" or choice =="a":
        dispatchReport()
    elif choice == "B" or choice =="b":
        fromBuilt()
    elif choice == "C" or choice =="c":
        appendBuild()
    elif choice=="D" or choice=="d":
        savePowersheet()
    elif choice=="E" or choice=="e":
        robPeter()
    elif choice=="F" or choice=="f":
        payPaul()
    elif choice=="Q" or choice=="q":
        sys.exit
    else:
        print("You must only select either A,B,C,D,E,F or Q.")
        print("Please try again")
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
        emptyCOl2 = p

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
        elif x is not None:
            print(x[:3],z)

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

    print('How many units will be needed to replace robbed power?')
    needed=input('If unknown, enter the amount of units stolen (NEED X).  ')
    
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
    print('Repaid check')

    dispatchReport()

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
                a = y[:3]
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
    menu()

def savePowersheet():
    print('Saving workboot.....')
    #wb.save('/home/'+getpass.getuser()+'/Documents/test.xlsx')
    wb.template = False
    wb.save('test.xlsx')
    menu()

if __name__ == '__main__':
    menu()
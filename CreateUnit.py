from py3270 import Emulator
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime, date 
# use x3270 so you can see what is going on
#loading and defining excel should be parameterized or stated in Prog.
wb1 = load_workbook(r'C:\Users\garren-james\Python Projects\AMS_Stress\AMS_User_IDs.xlsx')
ws1 = wb1['RACF User ID']
ws2 = wb1['Admins']
x = 1
y = input('Enter a number of records to process: ')
y= int(y)
rn = input('Enter row number to start from: ')
rn = int(rn)

now = datetime.now()
today = now.strftime("%m%d%Y")
#user details could be prompted for for, or included as kwargs if changed to a function.
usrId = 'fz0009'
pwrd = input('Enter the password for ' + usrId +":")
#Emulator set to run in background, can change to True to debug if errors
em = Emulator(visible=False)
em.connect('dcft.dcf.state.fl.us')
em.wait_for_field()
em.send_string('flors2')
em.send_enter()
em.wait_for_field()
em.fill_field(11, 36, usrId, 7)
em.fill_field(12, 36, pwrd, 8)
em.send_enter()
em.send_enter()
em.wait_for_field()

#create Unit
while x <= y:
    #define variable set from Excel File
    em.fill_field(23, 13, 'CLSM', 4)
    admin = 'C' + str(rn)
    admin = ws2[admin].value
    admin = str(admin)
    unit = admin
    dist = 'B' + str(rn)
    dist = ws2[dist].value
    dist = str(dist)
    newId = 'A' + str(rn)
    newId = ws2[newId].value
    newId = str(newId)
    em.fill_field(23, 29, admin + '/' + unit + '/' + dist, 14)
    em.send_enter()

    #Update the screen, when correctly loaded  
    if (em.string_found(1, 2, 'CLSM') == True):
        em.fill_field(11, 12,'AMS STRESS UNIT', 15)
        wrkNum = newId[-4:]
        wrkNum = int(wrkNum)+1
        wrkNum = str(wrkNum)
        if (len(wrkNum)==3):
            wrkNum = '0'+ wrkNum
        elif (len(wrkNum)==2):
            wrkNum = '00'+ wrkNum
        elif (len(wrkNum)==1):
            wrkNum = '000'+ wrkNum
        else:
            pass
        em.fill_field(11, 33, 'MS' + wrkNum, 6)
        em.fill_field(11, 68, today, 8)
        #select correct Service Site info based on Dist, would need to expand for more utility.
        if (dist=='10'):
            em.fill_field(12, 22, '06', 2)
            em.fill_field(12, 25, '400', 3)
        elif (dist=='11'):
            em.fill_field(12, 22, '13', 2)
            em.fill_field(12, 25, '401', 3)
        else:
            pass

    em.send_enter()

    #press Shift+F2 to refresh out to primary menu
    em.exec_command(b'PF(14)')
    em.exec_command(b'PF(14)')
    #iteration controls
    x += 1
    rn += 1

#simply display of status to catch where errors happen, or cycle batches
print(rn)
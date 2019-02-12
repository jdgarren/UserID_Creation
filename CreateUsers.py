from py3270 import Emulator
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime, date 
# use x3270 so you can see what is going on
wb1 = load_workbook(r'C:\Users\garren-james\Python Projects\UserID_Creation_Process\AMS_User_IDs.xlsx')
ws1 = wb1['RACF User ID']
ws2 = wb1['Admins']
x = 1
y = input('Enter a number of records to process: ')
y= int(y)
rn = input('Enter row number to start from: ')
rn = int(rn)
now = datetime.now()
today = now.strftime("%m%d%Y")
usrId = 'fz0009'
pwrd = input('Enter the password for ' + usrId +":")
ssnStart=100000000
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

while x <= y:
    startNewId = 'A' + str(rn)
    startDist = 'B' + str(rn)
    startAdmin = 'C' + str(rn)
    startTypeInd = 'D' + str(rn)
    newId = ws1[startNewId].value
    newId = str(newId)
    dist = ws1[startDist].value
    dist = str(dist)
    admin = ws1[startAdmin].value
    admin = str(admin)
    unit = str(admin)
    typeInd = ws1[startTypeInd].value
    typeInd = str(typeInd)
    em.wait_for_field()
    em.fill_field(23, 13, 'SMUM', 4)
    em.fill_field(23, 29, newId, 6)
    em.send_enter()
    em.wait_for_field()
    em.fill_field(3, 31, 'PC', 2)
    em.fill_field(3, 52, 'Y', 1)
    em.fill_field(3, 68, '0000', 4)
    em.fill_field(4, 19, 'STRESS TEST', 11)
    em.fill_field(4, 42, 'AMS', 3)
    em.fill_field(6, 7, str(ssnStart+rn), 9)
    em.fill_field(6, 23, dist, 2)
    if (dist=='11'):
        em.fill_field(6, 46, '13', 2)
        em.fill_field(6, 49, '401', 3)
    elif (dist=='10'):
        em.fill_field(6, 46, '06', 2)
        em.fill_field(6, 49, '400', 3)
    else:
        pass
    em.fill_field(6, 62, 'A', 1)
    if (typeInd=='a' or typeInd=='A'):
        em.fill_field(6, 78, '98', 2)
        em.fill_field(12, 24, 'OPA', 3)
    elif(typeInd=='s' or typeInd=='S'):
        em.fill_field(6, 78, '75', 2)
        em.fill_field(12, 24, 'ELIGPASS', 8)
    else:
        em.fill_field(6, 78, '50', 2)
        em.fill_field(12, 24, 'CASEPAS', 8)
    em.fill_field(12, 47, today, 8)
    em.send_enter()
    em.exec_command(b'PF(14)')
    em.exec_command(b'PF(14)')
    status = 'E'+str(rn)
    ws1[status] = "Done"
    x += 1
    rn += 1


wb1.save('AMS_User_IDs.xlsx')
print(rn)


from py3270 import Emulator
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime, date 
# use x3270 so you can see what is going on
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
usrId = 'fz0009'
pwrd = 'smile123'
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

#Build out Admin related variables here...
while x <= y:
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
    #create Admin
    em.fill_field(23, 13, 'CLAM', 4)
    em.fill_field(23, 29, admin + '/' + dist, 8)
    em.send_enter()
    em.fill_field(9, 12, 'AMS STRESS TEST', 15)
    em.fill_field(9, 32, newId, 6)
    em.fill_field(9, 70, today, 8)
    if (dist=='10'):
        em.fill_field(10, 32, '06', 2)
        em.fill_field(10, 35, '400', 3)
    elif (dist=='11'):
        em.fill_field(10, 32, '13', 2)
        em.fill_field(10, 35, '401', 3)
    else:
        pass

    em.send_enter()
    em.exec_command(b'PF(14)')
    em.exec_command(b'PF(14)')
    x += 1
    rn += 1

print(rn)
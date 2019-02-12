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

#start building caseloads for the user
while x <= y:

    em.fill_field(23, 13, 'CLCM', 4)
    admin = 'C' + str(rn)
    admin = ws1[admin].value
    admin = str(admin)
    unit = admin
    dist = 'B' + str(rn)
    dist = ws1[dist].value
    dist = str(dist)
    newId = 'A' + str(rn)
    newId = ws1[newId].value
    newId = str(newId)
    csldStrt = 0000
    csld = csldStrt+rn
    csld = str(csld)
    if (len(csld)==3):
        csld = '0'+ csld
    elif (len(csld)==2):
        csld = '00'+ csld
    elif (len(csld)==1):
        csld = '000'+ csld
    else:
        pass

    em.fill_field(23, 29, admin + '/' + unit + '/' + dist + csld +'/'+ newId, 25)
    em.send_enter()
    em.fill_field(8, 47, 'AMS STRESS CSLD', 15)
    em.fill_field(8, 78, 'N', 1)
    em.fill_field(12, 62, today, 8)
    em.send_enter()
    #add looping controls
    em.exec_command(b'PF(14)')
    em.exec_command(b'PF(14)')
    x += 1
    rn += 1
    print(rn)


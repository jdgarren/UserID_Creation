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
currPwrd = 'dcfams03'
newPwrd = 'dcfams04'


while x <= y:
    em = Emulator(visible=False)
    em.connect('dcft.dcf.state.fl.us')
    usrId = 'A' + str(rn)
    usrId = ws1[usrId].value
    usrId = str(usrId)
    usrId = 'A'+ usrId
    em.wait_for_field()
    em.send_string('flors2')
    em.send_enter()
    em.wait_for_field()
    em.fill_field(11, 36, usrId, 7)
    em.fill_field(12, 36, currPwrd, 8)
    em.fill_field(13, 61, newPwrd, 8)
    em.fill_field(14, 61, newPwrd, 8)
    em.send_enter()
    em.send_enter()
    em.send_enter()
    em.exec_command(b'PF(13)')
    em.send_enter()
    em.terminate()
    x += 1
    rn += 1
    print(rn)



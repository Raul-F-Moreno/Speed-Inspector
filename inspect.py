import openpyxl
import speedtest
from datetime import datetime
print("----Speed inspection begins----")
wb=openpyxl.load_workbook('work.xlsx')

st = speedtest.Speedtest()

sheet = wb.active

column = 'A'
cell = 1
location = f'{column}{cell}'

# Reads cell value, if cell stores none then goes to next.
def scanner():
    global cell
    global location
    place = sheet[location]
    read = [place.value]
    while True:
        if read != None:
            cell += 1
            location = f'{column}{cell}'
            place = sheet[location]
            read = place.value
        if read == None:
            return (location)
            break


now = datetime.now()
date = now.strftime("%d/%m/%Y %H:%M")
download = "{:.2f}".format(st.download()/1048576)
upload = "{:.2f}".format(st.upload()/1048576)
ping = f'{st.results.ping}'

print(f"Current date is {date}\n\
Current Download Speed is {download}\n\
Current Upload Speed is {upload}\n\
Ping is {st.results.ping}")

# Datetime information saves here.
dat = f'{scanner()}'
sheet[dat] = date

# Download Speed information saves here.
column = 'B'
cell = 1
cdo = f'{scanner()}'
sheet[cdo] = download

# Upload Speed information saves here.
column = 'C'
cell = 1
cup = f'{scanner()}'
sheet[cup] = upload

#Ping information saves here.
column = 'D'
cell = 1
cpi=f'{scanner()}'
sheet[cpi] = ping



print ('----Information has been saved-----')




wb.save('work.xlsx')

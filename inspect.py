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

#Information is saved
print('----Information has been saved-----\n')
wb.save('work.xlsx')

#Compares ping and establish jitter
def compare():
    column = 'D'
    cell = 1
    location = f'{column}{cell}'
    place = sheet[location]
    read = [place.value]
    while True:
        if read != None:
            cell += 1
            location = f'{column}{cell}'
            place = sheet[location]
            read = place.value
        if read == None:
            last_png = f'{column}{cell - 1}'
            previous_png = f'{column}{cell - 2}'
            before_png = f'{column}{cell - 3}'
            ping_a = sheet[last_png]
            ping_b = sheet[previous_png]
            ping_c = sheet[before_png]
            avg=((float(ping_a.value) + float(ping_b.value) + float(ping_c.value)/2))
            diff_1=(abs(float(ping_c.value) - float(ping_b.value)))
            diff_2 =(abs(float(ping_b.value)- float(ping_a.value)))
            avg_diff = int((diff_1 + diff_2)/2)
            jitter = (avg_diff/avg) * 100
            return(jitter)
            break

result = (float("{:.2f}".format(compare())))
print("----Stability Result----")
if result >= 15.0:
    print(f"Jitter is {result}% this connection might be unreliable.")
else:
    print(f"Jitter is {result}% Everything seems fine. For now.")

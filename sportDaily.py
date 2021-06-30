import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.mime.multipart import MIMEMultipart
import pandas
import xlsxwriter
workbook = xlsxwriter.Workbook('sportDaily.xlsx')

df = pandas.read_excel('./sportDaily.xlsx')
oldList = df.to_dict('records')
print(oldList)

day = int(input('day\t'))

def timeWork():
    h1 = int(s.split('.')[0])
    m1 = int(s.split('.')[1])
    h2 = int(e.split('.')[0])
    m2 = int(e.split('.')[1])    
    d = (h2*60+m2)-(h1*60+m1)
    hd = str(d // 60)
    md = str(d % 60)
    timeWork = float(hd+'.'+md)
    return timeWork
    
s = input('start\t')
start = float(s)
e = input('end\t')
end = float(e)
work = input('work\t')
tw = timeWork()

newList = {'DAY':day, 'START': start, 'END': end, 'WORK': work, 'TIME': timeWork}
oldList.append(newList)
list = oldList
print(list)

worksheet = workbook.add_worksheet()
headings = ['DAY', 'START', 'END', 'WORK', 'TIME']
worksheet.write_row('A1', headings)


for numDay in range(len(list)):
    for line in range(5):
        if line==0:
            worksheet.write(list[numDay]['DAY'], line, list[numDay]['DAY'])
        if line==1:
            worksheet.write(list[numDay]['DAY'], line, list[numDay]['START'])
        elif line==2:
            worksheet.write(list[numDay]['DAY'], line, list[numDay]['END'])
        elif line==3:
            worksheet.write(list[numDay]['DAY'], line, list[numDay]['WORK'])
        elif line==4:
            worksheet.write(list[numDay]['DAY'], line, list[numDay]['TIME'])

chart1 = workbook.add_chart({'type': 'bar'})
chart1.add_series({
	'name':	 ['Sheet1', 0, 4],
	'categories': ['Sheet1', 1, 0, numDay+1, 0],
	'values':	 ['Sheet1', 1, 4, numDay+1, 4],
})
chart1.set_style(11)
worksheet.insert_chart('G2', chart1)

workbook.close()

sender = 'send.by.py@gmail.com'
receivers = ['ebrahimimohammadali76@gmail.com']
body_of_email = 'Your daily data'

msg = MIMEMultipart()
msg['Subject'] = '=BACK UP=(Days:'+str(numDay+1)+')'
msg['From'] = sender
msg['To'] = ','.join(receivers)

part = MIMEBase('application', 'octet-stream')
part.set_payload(open('C:/Users/EBRAHIMI/Desktop/python prj/final to do/sportDaily.xlsx', 'rb').read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename = "sportDaily.xlsx"')
msg.attach(part)

s = smtplib.SMTP_SSL(host = 'smtp.gmail.com', port = 465)
s.login(user = sender, password = '@Lijun8080')
s.sendmail(sender, receivers, msg.as_string())
s.quit()

import getpass
import smtplib
import openpyxl as xl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# username = str(input('Your Username: '))
# password = str(input('Your Password: '))

username = str(input('Your email: '))
password = getpass.getpass(prompt='Your password: ')

From = username
Subject = 'Sample Subject'

wb = xl.load_workbook('./email_list.xlsx')
# sheet1 = wb.get_sheet_by_name('Sheet1') - this one is deprecated
sheet1 = wb['Sheet1']

names = []
emails = []

for cell in sheet1['A']:
    emails.append(cell.value)

for cell in sheet1['B']:
    names.append(cell.value)

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(username, password)

# For every person listed in the xslx, we are sending email

for i in range(len(emails)):
    msg = MIMEMultipart()
    msg['From'] = username
    msg['To'] = names[i]
    msg['Subject'] = Subject
    text = '''
Dear {},

Greetings. This is Hahn! Nice to meet you.

Have a great day.

Hahn
    '''.format(names[i])
    msg.attach(MIMEText(text, 'plain'))
    message = msg.as_string()
    server.sendmail(username, emails[i], message)
    print('Mail sent to', emails[i])

# When done, server quits

server.quit()
print('All emails sent successfully.')

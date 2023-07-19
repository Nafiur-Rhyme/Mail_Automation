'''
import csv
from time import sleep
import win32com.client as client

template = "{}, Wish you a Happy Birthday"

with open('people.csv', 'r', newline='') as f:
    reader = csv.reader(f)
    distro = [row for row in reader]

chunks = [distro[x:x+30] for x in range(0, len(distro), 30)]
outlook = client.Dispatch('Outlook.Application')

for chunk in chunks:
    for name, email in chunk:
        message = outlook.CreateItem(0)
        message.To = email
        message.Subject = "Happy Birthday"
        message.Body = template.format(name)
        message.Send()
    sleep(61)
'''

import csv
from time import sleep
import datetime
import win32com.client as client
import pathlib

card_path = pathlib.Path('Birthday_card.png')
card_abs = str(card_path.absolute())

with open('people.csv', 'r', newline='') as f:
    reader = csv.reader(f)
    recipients = [row for row in reader]
# Outlook only allows to send 30 mails per minute. That is why, chunks of 30 people are made
chunks = [recipients[x:x+30] for x in range(0, len(recipients), 30)]


#Function to check if today is birthday
def is_birthday(birthday):
    today = datetime.datetime.now().date()
    return today.month == birthday.month and today.day == birthday.day


#Function to send the email
def send_birthday_email():
    outlook = client.Dispatch('Outlook.Application')
    for chunk in chunks:
        for name, birthday, email in chunk:
            birthday = datetime.datetime.strptime(birthday, '%Y-%m-%d')
            if is_birthday(birthday):
                    outlook = client.Dispatch('Outlook.Application')
                    message = outlook.CreateItem(0)
                    img = message.Attachments.Add(card_abs)
                    html_body = f"""
                        <p>Dear {name},</p>
                        <div>
                            <img src="cid:card-img" width=50%>
                        </div>
                        """
                    img.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "card-img")
                    message.To = email
                    message.Subject = "Happy Birthday"
                    message.HTMLBody = html_body
                    message.Send()
        sleep(61)
        
send_birthday_email()
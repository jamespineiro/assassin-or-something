import win32com.client as win32

import csv

# ONLY WORKS ON WINDOWS WITH OUTLOOK RIP

outlook = win32.Dispatch('outlook.application')

names = []

emails = []

message = ""

 

with open("participants.csv","r")as file:

    reader = csv.DictReader(file)

    for row in reader:

        row["First Name"] = (row["First Name"].strip() + ' ' + row["Last Name"].strip()).title()

        new_name = row["First Name"]

        new_email = row["Email Address"]

        names.append(new_name)

        emails.append(new_email)

 

# for i in range(len(names)):

for i in range(1):

    message = ""

    mail = outlook.CreateItem(0)

    mail.Subject = 'Senior Assassin - Your target is...'

    mail.To = 'jpineiro2023@student.andoverma.us'

    #mail.To = emails[i]

    if i != len(names)-1: #emails[i]

        message+=f'Thank you for participating in Andover High School\'s 2023 Senior Assassin. <br><br>Your target is {names[i+1]}.'\

        '<br><br>---------------------------------------------<br>Senior Assassin Rules<br>---------------------------------------------<br>'\

        '(Rules here)<br>---------------------------------------------<br><br>This message was sent by an automation system created by James Pineiro,'\

        ' who is not associated with the Senior Assassin Instagram account. If you have any questions regarding the game or this email,'\

        ' please DM @ahs_seniorassassin2023 on Instagram.'

    else:

        message+=f'Your target is {names[0]}.<br><br>'

    mail.HTMLBody = (f"Dear {names[i]},<br><br>{message}")

    mail.Send()
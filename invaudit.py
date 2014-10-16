#!/usr/bin/python

import gspread
import sys
import smtplib
from email.mime.text import MIMEText
from tabulate import tabulate


# Function to send email
def sendmail(user_email, mytext):
    
    sender = 'sender@email.com'
    recipient = user_email
    subject = 'test email'

    headers = ["From: " + sender,
               "Subject: " + subject,
               "To: " + recipient,
               "mime-version: 1.0",
               "content-type: text/html"]
    headers = "\r\n".join(headers)
    headers += "\r\n\r\n"
    
    session = smtplib.SMTP('smtp.gmail.com:587')
    
    session.ehlo()
    session.starttls()
    session.ehlo
    session.login("sender@email.com", "senderpassword")
    
    session.sendmail(sender, recipient, headers + mytext)
    session.quit()
               
    print "Sent successfully"
    return


DEBUG=True

# Login with your Google account
gc = gspread.login('you@email.com', 'yourpassword')

# Open Spreadsheet1
if DEBUG:
    print 'Reading spreadsheet1...'
mylist1 = gc.open("Spreadsheet1").sheet1
# Get the names column
namelist = mylist1.col_values(1)
email_list = mylist1.col_values(5)

name_to_email ={}
i=0
while i < len (namelist):
    name_to_email[namelist[i]] = email_list[i]
    i+=1


# Exclude list
include_list = ['Jody',
                'John']

exclude_list = ['Caitlin', 'Andrew']

# Open Spreadsheet2
if DEBUG:
    print 'Reading spreadsheet2...'
newlist = gc.open("Spreadsheet2").sheet1
if DEBUG:
    print 'Reading spreadsheet2 - users column'
user_list = newlist.col_values(6)


# Dictionaries
devices={}
models={}
serials={}
users={}

for i in range(1, len(tag_list)):
    devices[tag_list[i]] = device_list[i]
    models[tag_list[i]] = model_list[i]
    serials[tag_list[i]] = sn_list[i]
    users[tag_list[i]] = user_list[i]


# Loop thru Spreadsheet1
for name1 in namelist:
    
    # Skip these names
    if name1 in exclude_list: continue
    
    # Only run names in include list
    # if name1 not in include_list: continue
    
    the_names = []
    the_tags = []
    the_devices = []
    the_models = []
    the_serials = []
    
    # Get equipment and details for the user
    for tag, name2 in users.items():
        if name1 == name2:
            the_names.append(name1)
            the_tags.append(tag)
            the_devices.append(devices[tag])
            the_models.append(models[tag])
            the_serials.append(serials[tag])

    # Create table
    mytable = tabulate({"Name": the_names,
                        "WMF Tag": the_tags,
                        "Device Type": the_devices,
                        "Model": the_models,
                        "Serial #": the_serials},
                        headers="keys")

    print mytable

    mytext = "Hi,<p>\n\nInsert text here.<p>\n\n"

    mytext += '<pre>\n'
    mytext += mytable
    mytext += '</pre>\n'

    mytext += "<p>\n\nMore text here<p>"


    # Send email
    if DEBUG:
        print '\nSending email to:', name_to_email[name1]
    sendmail(name_to_email[name1],mytext)



import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import smtplib 


scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']# API ACTIVATION

credentials = ServiceAccountCredentials.from_json_keyfile_name('D:\All Docs\SWMM SCHOL\InstanTeach\The Instanteach Assitant\Instanteach Assitant-95b1a9576d17.json',scope)#creditentials from Google to access Sheets

auth = gspread.authorize(credentials)#authenticate

response_sheet=auth.open('Instanteach: English Material Generator ').sheet1

last_row= 931
under_last_row=last_row+1
new_email=str(response_sheet.cell(under_last_row,3).value)
all_email_addresses=str(response_sheet.col_values(3))
previous_entry_checol_counter= all_email_addresses.count(new_email)




def main():
    last_row= 931
    under_last_row=last_row+1
    new_email=str(response_sheet.cell(under_last_row,3).value)
    all_email_addresses=str(response_sheet.col_values(3))
    previous_entry_checol_counter= all_email_addresses.count(new_email)
    print(previous_entry_checol_counter)

    last_row=931
    if previous_entry_checol_counter>1 :
        email_exists=True
    else:
        email_exists=False

    if '@' in new_email and email_exists==False:
        send_welcome_email()
        last_row=last_row+1
 

def send_welcome_email():
    teachername= str(response_sheet.cell(under_last_row,4).value)#teacher name
    fromaddr = 'instanteach.io@gmail.com'
    toaddrs  = new_email
    msg = ("Hi there dear {}!\n \n Thanks for using Instanteach! We have received your submission and we are working on preparing your material for you! \n \n You can now relax and sit back and rest assured that you'll receive all the material you asked for in a single email in less than 24hours.\n \nHere are some things that you can do instead of preparing your classes:\n \n -Watch your favourite series \n - Go outside and play some sport! ( being a teacher can be stressful, go get it all out) \n -Call a friend you haven't spoken to in a couple of month and see how they are doing\n \n* Did you know? From a study we did in 2018, teacher spent on average 4.5 hours preparing all their material forn their students. We are determined to end this and give this time back to you and so you can enjoy it again.\n \n We will be in touch with you shortly,\n \n Have a great time teaching!".format(teachername))
    username = 'instanteach.io@gmail.com'
    password = 'instanteachio88'
    server = smtplib.SMTP('smtp.gmail.com:587')
    server.ehlo()
    server.starttls()
    server.login(username,password)
    server.sendmail(fromaddr, toaddrs, msg)
    server.quit()


start=input('start?')
if start=='y':
    main()
    



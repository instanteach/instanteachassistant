# import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']# API ACTIVATION

credentials = ServiceAccountCredentials.from_json_keyfile_name('D:\All Docs\SWMM SCHOL\InstanTeach\The Instanteach Assitant\Instanteach Assitant-95b1a9576d17.json',scope)#creditentials from Google to access Sheets

auth = gspread.authorize(credentials)#authenticate

response_sheet=auth.open('Instanteach- Teacher Material Generator V1.2 (Responses)').sheet1

today_date = datetime.date.today() 


def done_checker(response_sheet):
   cell_notdone = response_sheet.find("Not Done")#Finds Row(teacher) Not Yet Done
   row_notdone = cell_notdone.row
   row_notdone_next=int((row_notdone)+1)
   name_notdone = str(response_sheet.cell(row_notdone,3).value)#teacher name
   timestamp_notdone = response_sheet.cell(row_notdone,1).value
   groupname_notdone =str(response_sheet.cell(row_notdone,4).value)
   number_notdone = response_sheet.cell(row_notdone,5).value
   agegroup_notdone = response_sheet.cell(row_notdone,6).value
   level_notdone = response_sheet.cell(row_notdone,7).value
   hoursneeded_notdone=response_sheet.cell(row_notdone,8).value
   email_notdone=str(response_sheet.cell(row_notdone,31).value)
   email_notdone_next=(response_sheet.cell(row_notdone_next,31).value)
   if email_notdone==email_notdone_next:
     print('watch out! Multiple entry!')
   os.chdir('D:\All Docs\SWMM SCHOL\InstanTeach\Instanteachers\Cohort B')
   exist_check=os.path.isdir(('D:\All Docs\SWMM SCHOL\InstanTeach\Instanteachers\Cohort B\{}').format(email_notdone))
   if exist_check==True:
     print('Teacher already exists ;)')
     os.startfile((('D:\All Docs\SWMM SCHOL\InstanTeach\Instanteachers\Cohort B\{}').format(email_notdone)))

   if exist_check==False:
    print('new teacher!')
    os.makedirs(('{}\{}').format(email_notdone,groupname_notdone))#createsfolders with teacher name and group name
    os.startfile(('D:\All Docs\SWMM SCHOL\InstanTeach\Instanteachers\Cohort B\{}\{}').format(email_notdone,groupname_notdone))
   print(('{} completed on {}').format(name_notdone,timestamp_notdone))
   print(('level:{}\n number of students:{}\n age group: {}\n lenght of class:{}').format(level_notdone,number_notdone,agegroup_notdone,hoursneeded_notdone))
   answered_cells = (response_sheet.cell(row_notdone,9).value),(response_sheet.cell(row_notdone,10).value),(response_sheet.cell(row_notdone,11).value),(response_sheet.cell(row_notdone,12).value),(response_sheet.cell(row_notdone,13).value),(response_sheet.cell(row_notdone,14).value),(response_sheet.cell(row_notdone,15).value),(response_sheet.cell(row_notdone,16).value),(response_sheet.cell(row_notdone,17).value),(response_sheet.cell(row_notdone,18).value),(response_sheet.cell(row_notdone,19).value),(response_sheet.cell(row_notdone,20).value),(response_sheet.cell(row_notdone,21).value),(response_sheet.cell(row_notdone,22).value),(response_sheet.cell(row_notdone,23).value),(response_sheet.cell(row_notdone,24).value),(response_sheet.cell(row_notdone,25).value),(response_sheet.cell(row_notdone,26).value),(response_sheet.cell(row_notdone,27).value),(response_sheet.cell(row_notdone,28).value),(response_sheet.cell(row_notdone,29).value),(response_sheet.cell(row_notdone,30).value)
   print(answered_cells)

   sendemail_prompt=input('Are you ready to send the email?(y/n)')
   if sendemail_prompt=='y' and exist_check==False:
     response_sheet.update_cell(row_notdone,2,'✓')
     response_sheet.update_cell(row_notdone_next,2,'Not Done')
     print(('{} Your Classroom Material For "{}" {}-Please Confirm That You Received This Email, Thanks!--\n \n Hi there {} !\n \n  I hope this email finds you very well! Please find attached  the material we suggest for your class, based on what you asked for in the questionnaire. I hope you find it useful!\n \n Remember that if you need more material, you can submit the questionnaire as many times as you need and for different groups: \n \n English Teaching Material Generator: https://goo.gl/forms/9iIFzmZKq0mSioT73  \n \n If you have any suggestions on how to improve the material or format, please don"t hesitate to tell me in an email! \n \n  Have a great time teaching! \n \n Conrad WS \n --Please Confirm That You Received This Email, Thanks!--').format(email_notdone,groupname_notdone,today_date,name_notdone))
   if sendemail_prompt=='y' and exist_check==True:
      response_sheet.update_cell(row_notdone,2,'✓')
      response_sheet.update_cell(row_notdone_next,2,'Not Done')
      print(('{} Your Classroom Material For "{}" {}-Please Confirm That You Received This Email, Thanks!--\n \n Hello again {} !\n Thanks for filling the form again! I appreciate your trust in us! \n \n Here is some more material based on what you filled out 😁!\n \n Let us know if this material was useful or not for your classes: \n \n And as always, feel free to fill out the questionaire as many times as you might need\n \n English Teaching Material Generator \n \n And have a great time teaching! \n \n Conrad WS').format(email_notdone,groupname_notdone,today_date,name_notdone))
   repeat_prompt=input('Repeat?(y/n)')
   if repeat_prompt=='y' or'Y':
     done_checker(response_sheet)
   if repeat_prompt=='n':
     quit()  
  
start_prompt=input('Hi, I m the Instanteach Assistant, here to accelerate the material sending to English teachers around the world.\n  Would you wish to start?(y/n)')

if start_prompt=='y' or 'Y':
  done_checker(response_sheet)


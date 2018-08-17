import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import datetime


scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']# API ACTIVATION

credentials = ServiceAccountCredentials.from_json_keyfile_name('D:\All Docs\SWMM SCHOL\InstanTeach\The Instanteach Assitant\Instanteach Assitant-95b1a9576d17.json',scope)#creditentials from Google to access Sheets

auth = gspread.authorize(credentials)#authenticate

response_sheet=auth.open('Instanteach: English Material Generator ').sheet1

today_date = datetime.date.today() 


def done_checker(response_sheet):
   cell_notdone = response_sheet.find("Not Done")#Finds Row(teacher) Not Yet Done
   row_notdone = cell_notdone.row
   row_notdone_next=int((row_notdone)+1)#locate row under
   name_notdone = str(response_sheet.cell(row_notdone,4).value)#teacher name
   timestamp_notdone = response_sheet.cell(row_notdone,1).value
   groupname_notdone =str(response_sheet.cell(row_notdone,5).value)
   number_notdone = response_sheet.cell(row_notdone,6).value
   agegroup_notdone = response_sheet.cell(row_notdone,7).value
   level_notdone = response_sheet.cell(row_notdone,8).value
   hoursneeded_notdone=response_sheet.cell(row_notdone,9).value
   email_notdone=str(response_sheet.cell(row_notdone,3).value)
   email_notdone_next=str(response_sheet.cell(row_notdone_next,3).value)
   Speaking_level=(response_sheet.cell(row_notdone,12).value)
   Writing_level=(response_sheet.cell(row_notdone,13).value)
   Listening_Level=(response_sheet.cell(row_notdone,14).value)
   Reading_level=(response_sheet.cell(row_notdone,15).value)
   Grammar_Level=(response_sheet.cell(row_notdone,16).value)
   Vocabulary_Level=(response_sheet.cell(row_notdone,17).value)
   exis_teacher_groupname=(response_sheet.cell(row_notdone,5).value),(today_date)
   if email_notdone==email_notdone_next:#checking for multiple entries by the same teacher
     print('watch out! Multiple entry!')
   os.chdir('D:\All Docs\SWMM SCHOL\InstanTeach\Instanteachers\Cohort C')
   exist_check=os.path.isdir(('D:\All Docs\SWMM SCHOL\InstanTeach\Instanteachers\Cohort C\{}').format(email_notdone))#Check if teacher previously filled out form
   if exist_check==True:
     print('Teacher already exists ;)')
     os.makedirs(('D:\All Docs\SWMM SCHOL\InstanTeach\Instanteachers\Cohort C\{}\{}').format(email_notdone,exis_teacher_groupname))#creates folders with teacher name and group name
     os.startfile(('D:\All Docs\SWMM SCHOL\InstanTeach\Instanteachers\Cohort C\{}\{}').format(email_notdone,exis_teacher_groupname))
     


   else :
      print('new teacher!')
      os.makedirs(('D:\All Docs\SWMM SCHOL\InstanTeach\Instanteachers\Cohort C\{}\{}').format(email_notdone,groupname_notdone))#creates folders with teacher name and group name
      os.startfile(('D:\All Docs\SWMM SCHOL\InstanTeach\Instanteachers\Cohort C\{}\{}').format(email_notdone,groupname_notdone))
   print(('{} completed on {} for {}').format(name_notdone,timestamp_notdone,groupname_notdone))
   print(('level:{}\n number of students:{}\n age group: {}\n lenght of class:{}').format(level_notdone,number_notdone,agegroup_notdone,hoursneeded_notdone))
   answered_cells = (response_sheet.cell(row_notdone,10).value),(response_sheet.cell(row_notdone,11).value),(response_sheet.cell(row_notdone,19).value),(response_sheet.cell(row_notdone,18).value)
   print(answered_cells) #Prints out the info inputted by the teacher
   print('Speaking:{}\n\n Writing:{}\n\n Listening:{}\n\nReading:{}\n\nGrammar:{}\n\n Vocab:{}\n\n'.format(Speaking_level,Writing_level,Listening_Level,Reading_level,Grammar_Level,Vocabulary_Level))
  
   sendemail_prompt=input('Are you ready to send the email?(y/n)')#Creates email template for new or returning teachers
   if sendemail_prompt=='y':
     if exist_check==False:
      response_sheet.update_cell(row_notdone,2,'✓')
      response_sheet.update_cell(row_notdone_next,2,'Not Done')
      print(('{}\n \n Your Classroom Material For "{}" {}\n \n ---Please Confirm That You Received This Email, Thanks!---\n \n Hi there {} !\n \n  I hope this email finds you very well! Please find attached  the material we suggest for your class, based on what you asked for in the questionnaire. I hope you find it useful!\n \n Remember that if you need more material, you can submit the questionnaire as many times as you need and for different groups: \n \n English Teaching Material Generator: http://instanteach.com/material-generator  \n \n If you have any suggestions on how to improve the material or format, please don"t hesitate to tell me in an email! \n \n ').format(email_notdone,groupname_notdone,today_date,name_notdone))
   if sendemail_prompt=='y' or sendemail_prompt== 'Y':
      if exist_check==True:
       response_sheet.update_cell(row_notdone,2,'✓')
       response_sheet.update_cell(row_notdone_next,2,'Not Done')
       print(('{}\n\n Your Classroom Material For "{}" {} \n\n Hello again {} !\n Thanks for filling the form again! I appreciate your trust in us! \n \n Here is some more material based on what you filled out 😁!\n \n Let us know if this material was useful or not for your classes: \n \n And as always, feel free to fill out the questionaire as many times as you might need: http://instanteach.com/material-generator\n \n English Teaching Material Generator \n \n').format(email_notdone,groupname_notdone,today_date,name_notdone))


   if sendemail_prompt=='y'or sendemail_prompt== 'Y':
      repeat_prompt=input('Repeat?(y/n)')#repeats for next entry 
      
   if repeat_prompt=='y' or'Y':
     done_checker(response_sheet)
   if repeat_prompt=='n':
     quit() and exit()  
  

def material_output():
    cell_notdone = response_sheet.find("Not Done")#Finds Row(teacher) Not Yet Done
    row_notdone = cell_notdone.row
    groupname_notdone =str(response_sheet.cell(row_notdone,5).value)
    email_notdone=response_sheet.cell(row_notdone,3).value
    output_material=os.listdir(os.startfile(('D:\All Docs\SWMM SCHOL\InstanTeach\Instanteachers\Cohort C\{}\{}').format(email_notdone,groupname_notdone)))
    response_sheet.update_cell(row_notdone,25,output_material)


start_prompt=input('Hi, I m the Instanteach Assistant, here to accelerate the material sending to English teachers around the world.\n  Would you wish to start?(y/n)')

if start_prompt=='y' or 'Y':#Start Prompt
  done_checker(response_sheet)

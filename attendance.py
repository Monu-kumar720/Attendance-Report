
# imported necessary modules
from http import server
import pandas as pd
import numpy as np
import os  
import datetime
import calendar
import csv 
from datetime import date
from datetime import datetime,timedelta
import itertools 
from csv import writer
from csv import DictWriter
from datetime import datetime
from datetime import date, timedelta
import smtplib
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.message import Message
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText
from email.message import EmailMessage
import ssl
start_time = datetime.now()

def attendance_report():
###Code
    print('hii')
    try:
        df1=pd.read_csv("input_attendance.csv")
        df2=pd.read_csv("input_registered_students.csv")
    except:
        print('error in input file')
        exit()

    df1=df1.dropna()# removing the null values 
    # splitting the columns
    #print(df1.head())
    df1[['Dates','Time']]=df1['Timestamp'].str.split(' ',expand=True)
    df1['Roll'] = df1.Attendance.str.split(' ', expand = True)[0]
    df1['Dates'] = pd.to_datetime(df1['Dates'],format = "%d-%m-%Y")
    #converting the date to day name
    df1['dayOfWeek'] = df1['Dates'].dt.day_name()
    # to calculate the unique dates in a list
    temp1=df1['Dates'].iloc[0]
    temp2=df1['Dates'].iloc[-1]
    final1 = temp1.strftime('%d-%m-%Y')
    final2=temp2.strftime('%d-%m-%Y')
    # x=temp1
    # datem1 = datetime.datetime.strptime(x, "%d-%m-%Y")
    # print(datem1)
    # print(type(final1))
    datem1 = datetime.strptime(final1, "%d-%m-%Y")
    # print(datem1)
    datem2 = datetime.strptime(final2, "%d-%m-%Y")
    s=date(datem1.year,datem1.month,datem1.day)
    e=date(datem2.year,datem2.month,datem2.day)
    delta=timedelta(days=1)
    li=[]#to store the monday and thrusday
    while(s<=e):
        x=s.strftime("%Y-%m-%d")
        xx=s.strftime("%Y-%m-%d")
        d = pd.Timestamp(xx)
        # print(d.dayofweek, d.day_name()) 
        # x_date = datetime.date(xx.)
        # no = x_date.weekday()
        
        if(d.day_of_week==0 or d.day_of_week==3):
           li.append(x)
        s+=delta
    # print((li))
    total_no_of_lectures=len(li)
        
    daTE=[]
    for i in df1['Dates']:
        
        daTE.append(i) 
    counti=0
    abc=[]
    for i in daTE:
        xy=(i.to_pydatetime())
        xyz=xy.strftime("%Y-%m-%d")
        abc.append(xyz)
         
    # print((daTE))  
    for i in li:
        for j in abc:
            if(i==j):
                counti+=1
     
    col=['Roll','Name']
    
    for i in li:
        col.append(i)
    col.append('Actual Lecture Taken')
    col.append('Total Real')
    col.append('%Attendence')
    #for storing all the attendence of students
    with open('file.csv','w')as file:
        csvWriter=csv.writer(file,delimiter=',')
        csvWriter.writerow(col)
        
        for (roll,name) in zip  (df2['Roll No'],df2['Name']):
            col1=[roll,name]
            count=0
            for datee in li:
                flag=0
                for (Roll,time,x) in zip(df1['Roll'],df1['Time'],abc):
                    if (Roll==roll and (time>='14:00:00'and time<='15:00:00')and x==datee):
                        col1.append('P')
                        flag=1
                        count+=1
                        break
                if flag==0:
                    col1.append("A")
                    
            col1.append(total_no_of_lectures)
            col1.append(count)
            col1.append(round(count/total_no_of_lectures*100,2))
            csvWriter.writerow(col1)
    df3=pd.read_csv('file.csv')
    resultExcelFile = pd.ExcelWriter('output\\attendance_report_consolidated.xlsx')#saving the file in excel 
    
    df3.to_excel(resultExcelFile,index=False)
    resultExcelFile.save()#saving the excel file
    os.remove('file.csv') #removing the file
    #for each students registred in the course
    for (roll,name) in zip  (df2['Roll No'],df2['Name']):
        col_index=['Date','Roll','Name','Total Attendance Count','Real','Duplicate','Invalid','Absent']
        col_index2=['',roll,name,'','','','','']
        with open('file1.csv','w')as file1:
            csvWriter=csv.writer(file1,delimiter=',')
            csvWriter.writerow(col_index)
            csvWriter.writerow(col_index2)
            
            for i in li:
                col_index3=[i]
                total_attendence=0
                real=0
                duplicate=0
                invalid=0
                absent=0
                curr=0
                for (Roll,time,x) in zip(df1['Roll'],df1['Time'],abc):
                    if (Roll==roll and (time>='14:00:00'and time<='15:00:00')and x==i):
                        curr+=1
                    elif (Roll==roll and (time<'14:00:00'and time>'15:00:00')and x==i):
                        invalid+=1
                if curr>=1:
                    real=1
                    duplicate=curr-real
                elif curr==0:
                    real=0
                    duplicate=curr-real
                    absent=1
                total_attendence=curr+invalid
                col_index3.append('')
                col_index3.append('')
                col_index3.append(total_attendence)
                col_index3.append(real)
                col_index3.append(duplicate)
                col_index3.append(invalid)        
                col_index3.append(absent)
                csvWriter.writerow(col_index3)
        
        df=pd.read_csv('file1.csv')
        resultexcelFile = pd.ExcelWriter('output\\'+roll+'.xlsx')##saving the file as student's roll no
        
        df.to_excel(resultexcelFile,index=False)
        resultexcelFile.save()
        os.remove('file1.csv')        
                        
    def send_email():                                                                           # Function to send email to cs3842022@gmail.com
        try:
            subject = "Consolidated Attendace Report"                                           
            body = "The report is attached with this mail."
            sender_email = input("Enter sender email : ")                                       # Sender e-mail
            receiver_email = "kumarmonuagrwl25@gmail.com"                                        # Receiver e-mail => cs3842022@gmail.com
            password = input("Type your password and press enter:")                             # Password of sender e-mail

            # Create a multipart message and set headers
            message = MIMEMultipart()                                                       
            message["From"] = sender_email
            message["To"] = receiver_email
            message["Subject"] = subject
            message["Bcc"] = receiver_email  # Recommended for mass emails

            # Add body to email
            message.attach(MIMEText(body, "plain"))

            filename = "output\\attendance_report_consolidated.xlsx"  # In same directory as script

            # Open csv file in binary mode
            with open(filename, "rb") as attachment:
                # Add file as application/octet-stream
                # Email client can usually download this automatically as attachment
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())

            # Encode file in ASCII characters to send by email    
            encoders.encode_base64(part)

            # Add header as key/value pair to attachment part
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {filename}",
            )
            # Add attachment to message and convert message to string
            message.attach(part)
            text = message.as_string()

            # Log in to server using secure context and send email
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(sender_email, password)
                server.sendmail(sender_email, receiver_email, text)
        except:
            print("Error in sending the file.")

    send_email()
        
    
        
            
            
     
    # df1.to_csv('output.csv',index=False)
    
    


    
    




from platform import python_version
ver = python_version()

if ver == "3.8.10":
    print("Correct Version Installed")
else:
    print("Please install 3.8.10. Instruction are present in the GitHub Repo/Webmail. Url: https://pastebin.com/nvibxmjw")


attendance_report()




#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))
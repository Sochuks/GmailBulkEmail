import pandas as pd
import smtplib


e = pd.read_excel("emailList.xlsx")
emails = e['Emails'].values

mailServer = smtplib.SMTP_SSL('smtp.gmail.com' , 465)

mailServer.login("adamconetwork@gmail.com", "rmauhwcwrcsrcvaa")

subject = input("Input Email Subject \n")
msg = input("Enter Email Message \n")
body = "Subject: {}\n\n{}".format(subject,msg)
# f'Subject: {subject}\n\n{msg}' Alternative format for above code

for email in emails:
    try:
        mailServer.sendmail("adamconetwork@gmail.com", email, body)
        print(" \n Message Delivered!")

    except:
        print("Email Failed to Deliver!")
        mailServer.quit()

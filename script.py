import pandas as pd
import smtplib

e = pd.read_excel("emailList.xlsx")
emails = e['Emails'].values

mailServer = smtplib.SMTP('smtp.gmail.com' , 587)
mailServer.starttls()
mailServer.login("adamconetwork@gmail.com", "rmauhwcwrcsrcvaa")

subject = input("Input Email Subject \n")
msg = input("Enter Email Message \n")
body = "Subject: {}\n\n{}".format(subject,msg)

for email in emails:
    try:
        mailServer.sendmail("adamconetwork@gmail.com", email, body)
        print(" \n Message Delivered!")

    except:
        print("Email Failed to Deliver!")
        mailServer.quit()

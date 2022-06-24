# Import required libraries
import pandas as pd
from datetime import date
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import random
from twilio.rest import Client
import getpass
import warnings

# To avoid warnings
warnings.filterwarnings("ignore")

# Gmail email login credentials -- Make sure the 2-step verification is on and supply App passwords here --
login_email = input("Enter Login Email: ") # email username
login_pass = getpass.getpass("Enter 2 step Verification App Password(Not Login Pass):") # Enter App passwords instead of Login Password

# Email configuration to send an email
s = smtplib.SMTP(host='smtp.gmail.com', port=587)
s.starttls()
s.login(login_email , login_pass)

# SMS config for sending SMS
# Account SID from twilio.com/console
account_sid = ""
# Auth Token from twilio.com/console
auth_token  = ""

def birthday_msg(name):
    with open('hbd.txt') as f:
        lines = f.readlines()
        msg1 = lines[0].replace("_", name)
        msg2 = random.choice(lines[1:])
        msg = msg1+"\n\n"+msg2
    return msg

def mrganv_msg(name):
    with open('mrgani.txt') as f:
        lines = f.readlines()
        msg1 = lines[0].replace("_", name)
        msg2 = random.choice(lines[1:])
        msg = msg1+"\n\n"+msg2
    return msg
        

def send_birthday_email(name, email):
    msg = MIMEMultipart()       # create a email message

    # Setup the email parameters
    msg['From']= login_email
    msg['Subject']="Happy Birthday"
    msg['To'] = email

    # The customized message to be emailed  
    message = birthday_msg(name)

    msg.attach(MIMEText(message, 'plain'))
    
    # sending the email
    s.send_message(msg)

    print ("Email Sent To:", msg['To'])

    del msg

def send_mrganv_email(name, email):
    msg = MIMEMultipart()       # create a message

    # Setup the email parameters
    msg['From']= login_email
    msg['Subject']="Happy Marriage Anniversary"
    msg['To'] = email

    # The customized message to be emailed 
    message = mrganv_msg(name)

    msg.attach(MIMEText(message, 'plain'))
    
    # sending the email
    s.send_message(msg)

    print ("Email Sent To:", msg['To'])

    del msg

def send_birthday_sms(name, mno):
    client = Client(account_sid, auth_token)
    sms = birthday_msg(name)
    message = client.messages.create(
        to=mno, 
        from_="+18573824350",
        body=sms)

    print("SMS sent To: ", mno)

def send_mrganv_sms(name, mno):
    client = Client(account_sid, auth_token)
    sms = mrganv_msg(name)
    message = client.messages.create(
        to=mno, 
        from_="+18573824350",
        body=sms)

    print("SMS sent To: ", mno)

def main():
    # Read the data file from its location     
    df = pd.read_excel (r"wish_data.xlsx")

    df = df.loc[lambda df: df['date']== date.today().strftime("%d/%m/%Y")]

    if not df.empty:
        # iterate through each row and select 
        for i in df.index:
            name = df['name'][i]
            mno = "+91"+str(df['mobile'][i])
            email = df['email'][i]

            if df['wish_type'][i] == 'birthday':
                send_birthday_email(name, email)
                send_birthday_sms(name,mno)
            elif df['wish_type'][i] == 'mrganv':
                send_mrganv_email(name, email)
                send_mrganv_sms(name,mno)
            else:
                print("No Wish Message Sent for ", name)
    else:
        print("There is no event Today")

    # close the smtp server 
    s.close()

if __name__ == "__main__":
    main()

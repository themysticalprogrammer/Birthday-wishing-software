import pandas as pd # Importing module to read excel files
from personal_data import * # Importing file having my personal data
import pywhatkit # Importing module that will help me send Whatsapp messages
import datetime # Importing module that will help me get current date
import smtplib # Importing Module that will help me to send E-mail
import os # Importing module that will help me to change directory

os.chdir(r"D:\Programming Projects\Python Projects\Birthday Wishing Software")

def sendEmail(to, sub, msg):
    smtplibVar = smtplib.SMTP('smtp.gmail.com', 587)
    smtplibVar.starttls()
    smtplibVar.login(GMAIL_ID, GMAIL_PASSWORD)
    smtplibVar.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    smtplibVar.quit()

def sendWhatsappMessage(number,message,hr,minutes):
    pywhatkit.sendwhatmsg(number,message, hr,minutes + 1)

if __name__ == '__main__':
    df = pd.read_excel('family_data.xlsx')

    currentDate = datetime.datetime.now().strftime("%d-%m")
    currentYear = datetime.datetime.now().strftime("%Y")

    writeInd = []

    for index, item in df.iterrows():
        bday = item['Birthday'].strftime("%d-%m")
        if(currentDate == bday) and currentYear not in str(item['Years_Wished']):
            sendEmail(item['Email_Address'],item['Subject'], item['Message'])
            currentTime = datetime.datetime.now()
            currentHour = currentTime.hour
            currentMinutes = currentTime.minute
            sendWhatsappMessage(item['Phone_Number'],item['Message'],currentHour,currentMinutes)
            writeInd.append(index)
        
    for i in writeInd:
        yr = df.loc[i, 'Years_Wished']
        df.loc[i, 'Years_Wished'] = str(yr) + ', ' + str(currentYear)

    df.to_excel('family_data.xlsx', index=False)
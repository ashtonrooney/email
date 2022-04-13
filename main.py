from random import randint
import pandas as pd
from email.message import EmailMessage
import smtplib
import logging
import time
import sys
import pdfkit
import os

path_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
config = pdfkit.configuration(wkhtmltopdf=path_wkhtmltopdf)
pdfkit.from_file('html_code.html', 'out.pdf', configuration=config)

logging.basicConfig(filename='mail.log', level=logging.DEBUG)

totalSend = 1
if(len(sys.argv) > 1):
    totalSend = int(sys.argv[1])

emaildf = pd.read_csv('gmail.csv')
contactsData = pd.read_csv('contacts.csv')  
bodies = ['bodies/body1.txt','bodies/body2.txt', 'bodies/body3.txt','bodies/body4.txt','bodies/body5.txt']
subjects = ['subjects/subject1.txt','subjects/subject2.txt', 'subjects/subject3.txt','subjects/subject4.txt','subjects/subject5.txt','subjects/subject6.txt','subjects/subject7.txt','subjects/subject8.txt','subjects/subject10.txt','subjects/subject11.txt']
froms = ['froms/from1.txt','froms/from2.txt','froms/from3.txt','froms/from4.txt','froms/from5.txt','froms/from6.txt','froms/from7.txt','froms/from8.txt','froms/from9.txt','froms/from10.txt','froms/from11.txt','froms/from12.txt']

def send_mail(name, email, emailId, password, bodyFile):
    newMessage = EmailMessage()

    # Invoice Number and Subject
    invoiceNo = randint(10000000, 99999999)
    l= randint(0,(len(subjects)-1))
    subject = open(subjects[l], 'r').read()
    subject = subject.replace('$name', name)
    subject = subject.replace('$invoice_no', str(invoiceNo))
    #subject = "Transaction Number -"  + str(invoiceNo) + " of item"
    

    newMessage['Subject'] = subject
    m= randint(0,(len(froms)-1))
    from1 = open(froms[m], 'r').read()
    newMessage['From'] = from1
    #newMessage['From'] = emailId
    newMessage['To'] = email
    transaction_id = randint(10000000, 99999999)

    # Mail Body Content
    body = open(bodyFile, 'r').read()
    body = body.replace('$name', name)
    body = body.replace('$invoice_no', str(transaction_id))

    # Mail PDF File
    html = open('html_code.html', 'r').read()
    html = html.replace('$name', str(name))
    html = html.replace('$email', str(email))
    html = html.replace('$invoice_no', str(transaction_id))
    # saving the changes to html_code.html
    with open('html_code.html', 'w') as f:
        f.write(html)

    file = "Invoice" + str(invoiceNo) + ".pdf"
    pdfkit.from_file('html_code.html', file, configuration=config)

    html = open('html_code.html', 'r').read()
    html = html.replace(str(transaction_id), '$invoice_no')
    with open('html_code.html', 'w') as f:
        f.write(html)

    newMessage.set_content(body)

    try:
        with open(file, 'rb') as f:
            file_data = f.read()
            file_name = f.name
        newMessage.add_attachment(
            file_data, maintype='application', subtype='octet-stream', filename=file_name)
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        #with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
            #smtp.starttls()
            smtp.login(emailId, password)
            smtp.send_message(newMessage)
            smtp.quit()

        os.remove(file)

        print(f"send to {email} by {emailId} successfully : {totalSend}")
        logging.info(
            f"send to {email} by {emailId} successfully : {totalSend}")

    except smtplib.SMTPResponseException as e:
        error_code = e.smtp_code
        error_message = e.smtp_error
        print(f"send to {email} by {emailId} failed")
        logging.info(f"send to {email}  by {emailId} failed")
        print(f"error code: {error_code}")
        print(f"error message: {error_message}")
        logging.info(f"error code: {error_code}")
        logging.info(f"error message: {error_message}")

        remove_email(emailId, password)


def start_mail_system():
    global totalSend
    j = 0
    k = 0

    for i in range(1,len(contactsData)):#change contact position to send.
        emaildf = pd.read_csv('gmail.csv')
        if(j >= len(emaildf)):
            j = 0
        time.sleep(0.03)
        send_mail(contactsData.iloc[i]['name'], contactsData.iloc[i]['email'], emaildf.iloc[j]['email'],
                  emaildf.iloc[j]['password'], bodies[k])
        totalSend += 1
        j = j + 1
        k = k + 1
        if j == 5:#change number of gmail to use
        #if j == len(emaildf):
            j = 0
        if k == len(bodies):
            k = 0
    quit()


def remove_email(emailId, password):
    df = pd.read_csv('gmail.csv')
    index = df[df['email'] == emailId].index
    df.drop(index, inplace=True)
    df.to_csv('gmail.csv', index=False)
    print(f"{emailId} removed from gmail.csv")
    logging.info(f"{emailId} removed from gmail.csv")
    


try:
    for i in range(6):
        start_mail_system()
except KeyboardInterrupt as e:
    print(f"\n\ncode stopped by user")

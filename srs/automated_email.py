#!/usr/bin/env python 3

import xlrd
import xlwt
import sqlite3
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

db = sqlite3.connect('C:\Users\Abdulaziz\Downloads\proto_database.sqlite')
cursor=db.cursor()


def save_excel_data_to_db():
    # Open the workbook and define the worksheet
    book = xlrd.open_workbook("C:\Users\Abdulaziz\Desktop\proto.xlsx")
    sheet = book.sheet_by_name("Sheet1")
    query = """INSERT INTO users (FIRSTNAME,LASTNAME,OTHERNAME,AGE) VALUES (?, ?, ?, ?)"""
    for r in range(1, sheet.nrows):
        firstName = sheet.cell(r, 0).value
        lastName = sheet.cell(r, 1).value
        otherName = sheet.cell(r, 2).value
        age = sheet.cell(r, 3).value

        #Assign values from each row
        values = (firstName, lastName, otherName, age)

        # Execute sql Query
        cursor.execute(query, values)

    # Close the cursor
    cursor.close()

    # Commit the transaction
    db.commit()

    # Close the database connection
    db.close()

def fetch_data_from_db():
    pass


def email_excel_data(username, password, emailTo, msg):
    message = MIMEMultipart()
    message['From'] = username
    message['To'] = emailTo
    message['Subject'] = 'A test mail sent by Python.'  # The subject line

    # The body and the attachments for the mail
    message.attach(MIMEText(msg, 'plain'))

    #Create SMTP session for sending the mail
    session = smtplib.SMTP('smtp.gmail.com', 587)#use gmail with port
    session.starttls()

    #enable security
    session.login(username, password) #login with mail_id and password
    text = message.as_string()
    session.sendmail(username, emailTo, text)
    session.quit()


def main():
    #save_excel_data_to_db()
    email_excel_data("abdulmoha786@gmail.com","ms@mbano786","abdulmoha786@gmail.com","Hello Zizu")


if __name__ == "__main__":
    main()

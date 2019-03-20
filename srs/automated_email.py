#!/usr/bin/env python 3
"""
This script This simple python script performs the following tasks
    -  Fetches data from the database
    - Opens an Excel Workbook
    - Creates a new Spreadsheet (Names the spreadsheet using the Time stamp DD-MM-YYYY H-M-S)
    - Appends the Data to the new sheet
    - Saves the workbook and emails it automatically to the specified recipients (in the program)
"""

import os
import xlrd
import xlwt
import sqlite3
import smtplib
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from mimetypes import guess_type
from email.encoders import encode_base64
from xlutils.copy import copy

# Edit this path to point to your location of the proto_database.sqlite file
# eg C:\Users\<Username>\Desktop\proto\proto_database.sqlite
db = sqlite3.connect(r'C:\Users\Abdulaziz\Downloads\proto_database.sqlite')


def save_excel_data_to_db():
    # Open the workbook and define the worksheet
    cursor = db.cursor()

    # Edit this path to point to your location of the proto.xls file
    # eg C:\Users\<Username>\Desktop\proto\proto.xls
    book = xlrd.open_workbook(r"C:\Users\Abdulaziz\Desktop\proto.xls")
    sheet = book.sheet_by_index(0)
    query = """INSERT INTO users (FIRSTNAME,LASTNAME,OTHERNAME,AGE) VALUES (?, ?, ?, ?)"""
    for r in range(1, sheet.nrows):
        # id = sheet.cell(r, 0).value
        firstName = sheet.cell(r, 0).value
        lastName = sheet.cell(r, 1).value
        otherName = sheet.cell(r, 2).value
        age = sheet.cell(r, 3).value

        # Assign values from each row
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
    cursor = db.cursor()
    cursor.execute("SELECT * FROM users")
    rowNum = 1  # keep track of rows

    # print all the cells of the row to excel sheet
    # Edit this path to point to your location of the proto.xls file
    # eg C:\Users\<Username>\Desktop\proto\proto.xls
    book = xlrd.open_workbook(r"C:\Users\Abdulaziz\Desktop\proto.xls")
    wb = copy(book)
    sheetName = datetime.datetime.today().strftime("%Y-%m-%d %H-%M-%S")
    sheet = wb.add_sheet(sheetName)
    #write column headers
    sheet.write(0, 0, "ID")
    sheet.write(0, 1, "FIRSTNAME")
    sheet.write(0, 2, "LASTNAME")
    sheet.write(0, 3, "OTHERNAME")
    sheet.write(0, 4, "AGE")

    for row in cursor.fetchall():
        colNum =0
        for col in row:
            sheet.write(rowNum, colNum, col)  # row, column, value
            colNum += 1
        rowNum = rowNum + 1

    # Edit this path to point to your location of the proto.xls file
    # eg C:\Users\<Username>\Desktop\proto\proto.xls
    wb.save(r"C:\Users\Abdulaziz\Desktop\proto.xls")
    cursor.close()


def email_excel_data(username, password, emailTo, msg, attachments):
    fetch_data_from_db()
    message = MIMEMultipart()
    message['From'] = username
    message['To'] = emailTo
    message['Subject'] = 'A test mail sent by Python.'  # The subject line

    # The body and the attachments for the mail
    message.attach(MIMEText(msg, 'plain'))
    if attachments is not None:
        for filename in attachments:
            mimetype, encoding = guess_type(filename)
            mimetype = mimetype.split('/', 1)
            fp = open(filename, 'rb')
            attachment = MIMEBase(mimetype[0], mimetype[1])
            attachment.set_payload(fp.read())
            fp.close()
            encode_base64(attachment)
            attachment.add_header('Content-Disposition', 'attachment',
                                  filename=os.path.basename(filename))
            message.attach(attachment)

    # Create SMTP session for sending the mail
    session = smtplib.SMTP('smtp.gmail.com', 587)  # use gmail with port
    session.starttls()

    # enable security
    session.login(username, password)  # login with mail_id and password
    text = message.as_string()
    session.sendmail(username, emailTo, text)
    session.quit()


def main():
    # Edit this path to point to your location of the proto.xls file
    # eg C:\Users\<Username>\Desktop\proto\proto.xls

    attachments = [r"C:\Users\Abdulaziz\Desktop\proto.xls"]

    # Edit the below methods arguments to match your credentials
    # 1 Your Username(Email address specifically Gmail)
    # 2 Your Password
    # 3 The Recipient Email Address(Include the same as yours for testing purposes)
    # 4 The Message to be included in the body
    email_excel_data("abdulmoha786@gmail.com", "*********", "abdulmoha786@gmail.com", "Hello Zizu, Here is your excel file", attachments)


if __name__ == "__main__":
    main()

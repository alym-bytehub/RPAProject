# Insert your python code here

import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

# INSERT ANSWER FOR QUESTION 1 HERE
directory = os.getcwd()+'\\'
# directory = "C:\Users\mccab\OneDrive\Documents\Advanced Analytics\17 - Project 3 - Robotic Process Automation\\"
while True:
    directory = single_input('Please input directory where files are located.')
    directory = directory + '\\'
    if directory.endswith('\\'):
        break
    else:
        directory = directory + '\\'
        if os.path.exists(directory + 'ProductData.xlsx' + '\\') == True and os.path.exists(directory + 'MasterInvoice.pdf' + '\\') == True and os.path.exists(directory + 'Order1.xlsx' + '\\') == True and os.path.exists(directory + 'Order2.xlsx' + '\\') == True and os.path.exists(directory + 'Order3.xlsx' + '\\') == True and os.path.exists(directory + 'Order4.xlsx' + '\\') == True:
            break
        else:
            show_message('Error', 'Directory does not contain all files. Please try again.')
print(directory)

# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}
product_desc = pd.read_excel(directory + '//' + 'ProductData.xlsx', sheet_name = 'Descriptions')
for i, row in product_desc.iterrows():
    descriptions = row['Description'] #value
    unit_number = row['Unit number'] #key
    product_descriptions[unit_number] = descriptions
product_prc = pd.read_excel(directory + '//' + 'ProductData.xlsx', sheet_name = 'Prices')
for i, row in product_prc.iterrows():
    prc_per_unit = row['Price per unit'] #value
    unit_number = row['Unit number'] #key
    product_prices[unit_number] = prc_per_unit
print(product_descriptions)
print(product_prices)

# INSERT ANSWER FOR QUESTION 3 HERE

def get_order_data(invoice_num, filename, directory):
    order_wb = openpyxl.load_workbook(directory + filename)
    order = order_wb['OrderForm']
    cust_name = order['C14'].value
    cust_email = order['C13'].value
    order_date = order['C9'].value
    company = order['C15'].value
    address = order['C16'].value
    city = order['C17'].value
    state = order['C18'].value
    zipcode = order['C19'].value
    order_date = order_date.strftime('%m/%d/%Y')

    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = str(order_date)
    order_data['CustomerAddress1'] = str(address)
    order_data['CustomerName'] = str(cust_name)
    order_data['CustomerCompany'] = str(company)
    order_data['CustomerAddress2'] = str(city) + ', ' + str(state) + ' ' + str(zipcode)

    order_total = 0
    item = 0
    rows = [24, 25, 26, 27, 28]

    for row in rows:
        unit_number = order['B' + str(row)].value
        quantity = order['C' + str(row)].value
        if unit_number is not None:
            description = product_descriptions[unit_number]
            price = product_prices[unit_number]
            amount = price * quantity

            order_data['Description' + str(item)] = str(description)
            order_data['Price' + str(item)] = str(price)
            order_data['Quantity' + str(item)] = str(quantity)
            order_data['Amount' + str(item)] = str(amount)

            order_total = order_total + amount
            item = item + 1

    order_data['Total'] = str(order_total)
    order_wb.close()
    return order_data, cust_name, cust_email, order_date
order_data, name, email, date = get_order_data(8515, 'Order1.xlsx', directory)
print(order_data)
print(name)
print(email)
print(date)

# INSERT ANSWER FOR QUESTION 4 HERE
def create_invoice(invoice_num, order_data, directory):
    filename = 'Invoice ' + str(invoice_num) + '.pdf'
    og_invoice_path = directory + 'MasterInvoice.pdf'
    new_invoice_path = directory + 'Invoice ' + str(invoice_num) + '.pdf'
    create_pdf(og_invoice_path, new_invoice_path, order_data)
create_invoice(8515, order_data, directory)

# INSERT ANSWER FOR QUESTION 5 HERE
def send_email(invoice_num, name, email, date, directory):
    gmail_user = "mccabe.alyssa.c@gmail.com"
    gmail_password_path = select_file('Please select the gmail password file')
    gmail_password = open(gmail_password_path).read()

    recipients = [email]
    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'
    body = f'''Dear {name}, 

    Attached is your invoice dated {date}. 

    Thank you for your business. 

    PythonTech Co. 
    10 Python Drive
    Pleasant Grove, UT 84602
    Phone: (801)555-2372
    Email: sales@pythontechco.com'''

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    path_to_file = directory + f'Invoice {invoice_num}.pdf'
    message = add_attachment(path_to_file, message)

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    mail_server.login(gmail_user, gmail_password)

    mail_server.send_message(message)
    mail_server.quit()

send_email(8515, 'Thomas Holt', 'amccabe1@byu.edu', '11/13/2023', directory)

# INSERT ANSWER FOR QUESTION 6 HERE
invoice_num = 8515
filenames = os.listdir(directory)
for filename in filenames: 
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1


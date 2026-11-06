# IMPORT PACKAGES IN

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

#1
while True:
    directory = single_input('Please input the directory to the folder where ProductData.xlsx, MaterInvoice.pdf, and Order Files are located.')
    directory = directory + '\\'

    try:
        file_names = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']
        directory_files = os.listdir(directory)

        if file_names[0] in directory_files and file_names[1] in directory_files and file_names[2] in directory_files and file_names[3] in directory_files and file_names[4] in directory_files and file_names[5] in directory_files:
            break
        else:
            show_message('Error','Please make sure you entered the correct folder path, with the necessary files in the folder.')
    except:
        show_message('Error','Please make sure you entered the correct folder path, with the necessary files in the folder.')    

print(directory)
#2
product_descriptions = {}
product_data = pd.read_excel(directory+'ProductData.xlsx',sheet_name='Descriptions')
for index, row in product_data.iterrows():
    unit_num = row['Unit number']
    descrip = row['Description']
    product_descriptions[unit_num] = descrip

product_prices = {}
product_data = pd.read_excel(directory+'ProductData.xlsx',sheet_name='Prices')
for index, row in product_data.iterrows():
    unit_num = row['Unit number']
    price = row['Price per unit']
    product_prices[unit_num] = price

#3
def get_order_data(invoice_num, filename, directory):
    order_data = {}
    order = openpyxl.load_workbook(filename)
    selected_sheet = order['OrderForm']
    name = selected_sheet['C14'].value
    email = selected_sheet['C13'].value
    date = selected_sheet['C9'].value

    unit0 = selected_sheet['B24'].value
    quant0 = selected_sheet['C24'].value
    price0 = product_prices.get(unit0)
    descript0 = product_descriptions.get(unit0)
    amount0 = quant0 * price0
    total = amount0

    city = selected_sheet['C17'].value
    state = selected_sheet['C18'].value
    zip = selected_sheet['C19'].value

    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date.strftime('%m/%d/%Y')
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress2'] = f'{city}, {state} {zip}'
    
    order_data['Price0'] = f'${price0:,.2f}'
    order_data['Quantity0'] = str(quant0)
    order_data['Amount0'] = f'${amount0:,.2f}'
    order_data['Description0'] = descript0

    unit1 = selected_sheet['B25'].value
    quant1 = selected_sheet['C25'].value 
    unit2 = selected_sheet['B26'].value
    quant2 = selected_sheet['C26'].value

    if unit1 is not None:
        price1 = product_prices.get(unit1)
        descript1 = product_descriptions.get(unit1)
        amount1 = quant1 * price1
        order_data['Price1'] = f'${price1:,.2f}'
        order_data['Quantity1'] = str(quant1)
        order_data['Amount1'] = f'${amount1:,.2f}'
        order_data['Description1'] = descript1
        total = total + amount1
    elif unit2 is not None:
        price2 = product_prices.get(unit2)
        descript2 = product_descriptions.get(unit2)
        amount2 = quant2 * price2
        order_data['Price2'] = f'${price2:,.2f}'
        order_data['Quantity2'] = str(quant2)
        order_data['Amount2'] = f'${amount2:,.2f}'
        order_data['Description2'] = descript2
        total = total + amount2

    order_data['Total'] = f'${total:,.2f}'

    return order_data, name, email, date

#4
def create_invoice(invioce_num,order_data,directory):
    create_pdf(directory+'MasterInvoice.pdf',directory+f'Invoice {invioce_num}.pdf', order_data)

#5
gmail_password_path = select_file('Please select the gmail password file.')

def send_email(invoice_num, name, email, date, directory):

    gmail_user = "eastonfost@gmail.com"
    gmail_password = open(gmail_password_path).read()

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    body = f"""
    Dear {name},

    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com
    """

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = email
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory+f'Invoice {invoice_num}.pdf'
    #path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()

#6
filenames = os.listdir(directory)
invoice_num = 8515

for name in filenames:
    if 'Order' in name and 'xlsx' in name:
        order_data, name, email, date = get_order_data(invoice_num, name, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num,name,email,date,directory)
        invoice_num = invoice_num + 1
    else:
        pass
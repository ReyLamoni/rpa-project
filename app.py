# IMPORT PACKAGES IN THIS CELL

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

# part 1
while True:
    directory = single_input('Please input the path to the folder where your ProductData.xlsx, MasterInvoice.pdf, and Order files are stored.', width= 500) # will the pixels help with the display? or do we do a ''' '''?
    directory = directory.strip('"') #get rid of the quotes that auto populate when you copy in the path
    if os.path.exists(directory):
        if directory[-1] != '\\':
            directory = directory+'\\'
        dir_check = os.listdir(directory)
        required_files = ['ProductData.xlsx','MasterInvoice.pdf','Order1.xlsx','Order2.xlsx','Order3.xlsx','Order4.xlsx'] #is there a better way?
        has_all_files = True
        for file in required_files:
            if file not in dir_check:
                has_all_files = False
        if has_all_files == True:
            break
        else:
            show_message('Error:','The folder you have selected does not contain the required files, please choose another folder that contains the following files: ProductData.xlsx, MasterInvoice.pdf, and Order1.xlsx, Order2.xlsx, Order3.xlsx, Order4.xlsx, ..., Order[].xlsx', width=600)
            
    else:
        show_message('Error:','Sorry, you selected a file path that does not exist. Please try again')

# part 2
# step 1 -> read in descriptions and create product descriptions dictionary
product_descriptions = {}
descriptions = pd.read_excel(directory+'ProductData.xlsx',sheet_name='Descriptions')
for i, row in descriptions.iterrows():
    key = row['Unit number']
    value = row['Description']
    product_descriptions[key] = value

# part 2 --> read in the price data and create a dictionary
product_prices = {}
descriptions = pd.read_excel(directory+'ProductData.xlsx',sheet_name='Prices')
for i, row in descriptions.iterrows():
    key2 = row['Unit number']
    value2 = row['Price per unit']
    product_prices[key2] = value2


# part 3
def get_order_data(invoice_num, filename, directory):

    template = openpyxl.load_workbook(directory+filename)
    order = template['OrderForm']

    name = order['C14'].value
    email = order['C13'].value
    raw_date = str(order['C9'].value)
    year = raw_date.split()[0].split('-')[0]
    month = raw_date.split()[0].split('-')[1]
    day = raw_date.split()[0].split('-')[2]
    date = f'{month}/{day}/{year}'
    
    update_dict = {}
    update_dict['InvoiceNum'] = str(invoice_num)
    update_dict['Date'] = date
    update_dict['CustomerName'] = name
    update_dict['CustomerAddress1'] = order['C16'].value
    update_dict['CustomerAddress2'] = f'{order['C17'].value}, {order['C18'].value} {order['C19'].value}'
    update_dict['CustomerCompany'] = order['C15'].value

    total = 0

    if order['B24'].value is not None:
        price = product_prices[order['B24'].value]
        quantity = order['C24'].value
        amount = price*quantity
        update_dict['Description0'] = product_descriptions[order['B24'].value]
        update_dict['Price0'] = f'${price:,.2f}'
        update_dict['Quantity0'] = str(quantity)
        update_dict['Amount0'] = f'${amount:,.2f}'
        total = total + amount

    if order['B25'].value is not None:
        price = product_prices[order['B25'].value]
        quantity = order['C25'].value
        amount = price*quantity
        update_dict['Description1'] = product_descriptions[order['B25'].value]
        update_dict['Price1'] = f'${price:,.2f}'
        update_dict['Quantity1'] = str(quantity)
        update_dict['Amount1'] = f'${amount:,.2f}'
        total = total + amount

    if order['B26'].value is not None:
        price = product_prices[order['B26'].value]
        quantity = order['C26'].value
        amount = price*quantity
        update_dict['Description2'] = product_descriptions[order['B26'].value]
        update_dict['Price2'] = f'${price:,.2f}'
        update_dict['Quantity2'] = str(quantity)
        update_dict['Amount2'] = f'${amount:,.2f}'
        total = total + amount

    update_dict['Total'] = f'${total:,.2f}'

    return update_dict, name, email, date


# PART 4
def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory+'MasterInvoice.pdf',directory+f'Invoice {invoice_num}.pdf',order_data)



#Part 5

def send_email(invoice_num, name, email, date, directory):
    # Enter your gmail user name and password
        #this section has been moved out of the function to not repeat everytime you send the email

    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    name = name
    
    body = f"""Dear {name},
        
    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory+f'Invoice {invoice_num}.pdf'
    
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


# adjusted part 5 to select the email
while True:
    gmail_user = single_input('Please input the gmail account you would like to send the email from:')
    if "@gmail.com" not in gmail_user:
        show_message('Please try again:','You did not enter a valid email address')
    else:
        break
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()


#PART 6
filenames = os.listdir(directory)
invoice_num = 8515
for filename in filenames:
    if 'Order' in filename and '.xlsx' in filename:
        order_data, name, email, date = get_order_data(invoice_num,filename,directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email,date, directory)
        invoice_num = invoice_num+1
# Insert your python code here

# # INSERT ANSWER FOR QUESTION 1 HERE

# invoice_data_path = select_file(title="Please select the data.")

# directory = invoice_data_path.split('/')
# del directory[-1]
# directory = '/'.join(directory)
# directory = directory + '//'
# print(directory)
inputs = ['ProductData.xlsx', 'MasterInvoice.pdf', 'Order1.xlsx', 'Order2.xlsx', 'Order3.xlsx', 'Order4.xlsx']

while True:
    invoice_data_path = select_file(title="Please select the Invoice-Data.xlsx file.")
    directory = invoice_data_path.split('/')
    del directory[-1]
    directory = '/'.join(directory)
    directory = directory + '/'
    try:
        filenames = os.listdir(directory)
        if ('ProductData.xlsx' in filenames) and ('MasterInvoice.pdf' in filenames) and ('Order1.xlsx' in filenames) and ('Order2.xlsx' in filenames) and ('Order3.xlsx' in filenames) and ('Order4.xlsx' in filenames):
            break
        else:
            show_message("Error", "Please try again.")

    except:
        show_message("Error", "Please try again.")


# INSERT ANSWER FOR QUESTION 2 HERE
product_descriptions = {}
product_prices = {}

df_descriptions = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
for index,row in df_descriptions.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

df_prices = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
for index,row in df_prices.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']


print(product_descriptions)
print(product_prices)

# # INSERT ANSWER FOR QUESTION 3 HERE
from datetime import datetime


def get_order_data(invoice_num, filename, directory):
    filepath = directory + filename
    wb = openpyxl.load_workbook(filepath, data_only=True)
    form = wb.worksheets[0] ##?
    name = str(form['C14'].value)
    email = str(form['C13'].value)
    date = form['C9'].value
    



# get_pdf('MasterInvoice.pdf')
    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = f'{date.month}/{date.day}/{date.year}'
    order_data['CustomerCompany'] = str(form['C15'].value)
    order_data['CustomerName'] = str(form['C14'].value)
    order_data['CustomerAddress1'] = str(form['C16'].value)
    city = form['C17'] .value
    state = form['C18'].value
    zip_code = form['C19'].value
    order_data['CustomerAddress2'] = f'{city}, {state}, {zip_code}'

    total = 0
    cell_num = 24
    product_num = 0
    while True:
        unitnum = form[f'B{cell_num}'].value
        if unitnum is None:
            break
        else:
            quantity = form[f'C{cell_num}'].value
            price = product_prices[unitnum]
            amount = price * quantity
            print(unitnum, price, quantity, amount, total)
            order_data[f'Description{product_num}'] = product_descriptions[unitnum]
            order_data[f'Price{product_num}'] = f'${price:,.2f}'
            order_data[f'Quantity{product_num}'] = f'{quantity:,.2f}'
            order_data[f'Amount{product_num}'] = f'${amount:,.2f}'
            total = total + amount
            cell_num = cell_num + 1
            product_num = product_num + 1
    order_data['Total'] = f'${total:,.2f}'
    return order_data, name, email, date


# INSERT ANSWER FOR QUESTION 4 HERE
# df1 = get_pdf('MasterInvoicetest.pdf')
# df1

def create_invoice(invoice_num, order_data, directory):
    template_path = directory + 'MasterInvoice.pdf'
    new_path = directory + f'Invoice {invoice_num}.pdf'
    create_pdf(template_path, new_path, order_data)

order_data = get_order_data(8515, 'Order1.xlsx', directory)
create_invoice(8515, order_data, directory)
print(order_data)

# INSERT ANSWER FOR QUESTION 5 HERE
from email.message import EmailMessage
gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):


    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    
    body = f"""Dear {name},

    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co.
    10 Python Drive
    Pleasant Grove, UT 84062
    Phone: (801) 555-2372
    Email: sales@pythontechco.com
    """
    message = EmailMessage()
    message['From'] = "PythonTech Co."
    message['To'] = email
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = directory + f'Invoice {invoice_num}.pdf'

    #path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


send_email(8515, 'Noah', 'noahnielson@gmail.com', '11/13/2023', directory)


# INSERT ANSWER FOR QUESTION 6 HERE
filenames = os.listdir(directory)
filenames
invoice_num = 8515
for file in filenames:
    if 'Order' in file and 'xlsx' in file: 
        order_data, name, email, date = get_order_data(invoice_num, file, directory)
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num = invoice_num + 1
    
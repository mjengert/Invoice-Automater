import openpyxl, os, win32printing, win32com.client as win32


# gets the last invoice number from the last invoice made and makes a new one
def new_invoice_num():
    invoice_num_data = []
    past_invoice_num = new_invoice_sheet['B1'].value
    for value in past_invoice_num:
        if value.isdigit():
            index = past_invoice_num.index(value)
            invoice_number = past_invoice_num[index:]
            invoice_num = int(invoice_number) + 1
    invoice_num_data.append(invoice_num)
    str_num = past_invoice_num[:index] + str(invoice_num)
    invoice_num_data.append(str_num)
    return invoice_num_data


filename = 'Invoice Excel Docs'
directory = os.path.abspath(filename)
invoice_data = {}
# open invoice template
for invoices in os.listdir(directory):
    if not invoices.endswith('.xlsx'):
        break
    invoice = openpyxl.load_workbook(os.path.join(directory, invoices))
    invoice_sheet = invoice.active
    invoice_number = invoice_sheet['B1'].value
    for value in invoice_number:
        if value.isdigit():
            index = invoice_number.index(value)
            invoice_number = invoice_number[index:]
            break
    invoice_num = int(invoice_number)
    name = invoice_sheet['B7'].value
    address = invoice_sheet['B8'].value
    phone_number = invoice_sheet['B9'].value
    invoice_data[name] = {'customer_name': name, 'address': address, 'phone_number': phone_number}


new_invoice = openpyxl.load_workbook('Invoice_template.xlsx')
new_invoice_sheet = new_invoice.active
Invoicemaker = True
while Invoicemaker:
    name = input('Enter customer name: ')
    if name in invoice_data:
        print('Customer name already exists')
        print('Making new invoice for customer from existing data')
        new_invoice_num = new_invoice_num()
        new_invoice_sheet['B1'] = new_invoice_num[1]
        new_invoice_sheet['B7'] = invoice_data[name]['customer_name']
        new_invoice_sheet['B8'] = invoice_data[name]['address']
        new_invoice_sheet['B9'] = invoice_data[name]['phone_number']
    else:
        print('Customer name does not exist')
        print('Making new invoice for new customer')
        new_invoice_sheet['B7'] = name
        address = input('Enter customer address: ')
        parts = address.split(", ")
        street = parts[0]
        city_state_zip = parts[1:]
        city, state, zip_code = city_state_zip[0], city_state_zip[1], city_state_zip[2]
        new_address = f"{street}\n{city}, {state.upper()} {zip_code}"
        new_invoice_sheet['B8'] = new_address
        new_invoice_num = new_invoice_num()
        new_invoice_sheet['B1'] = new_invoice_num[1]
        new_invoice_sheet['B9'] = input('Enter customer phone number: ')
        invoice_data[name] = {'customer_name': name, 'address': new_address, 'phone_number': new_invoice_sheet['B9'].value}

    new_invoice_sheet['C7'] = input('What is this invoice for? ')
    new_invoice_sheet['B13'] = input('Enter the details: ')
    new_invoice_sheet['C13'] = int(input('Enter the price: '))

    new_invoice_template = new_invoice
    new_invoice_template.save('Invoice_template.xlsx')
    new_invoice.save(f'{name} #{new_invoice_num[0]}.xlsx')
    print(f'{new_invoice_num[0]} has been saved!')

    # determines if user wants to make another invoice
    continue_maker = input('Would you like to make another invoice? (y/n) ')
    if continue_maker == 'y':
        continue
    else:
        print('Thank you for using the Invoice Maker!')
        Invoicemaker = False



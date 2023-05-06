import openpyxl, os, shutil, tkinter
from tkinter import *
from tkinter import font as tkfont

# gets the last invoice number from the last invoice made and makes a new one
def new_invoice_num():
    invoice_num_data = []
    past_invoice_num = new_invoice_sheet['B1'].value
    for num in past_invoice_num:
        if num.isdigit():
            index_v = past_invoice_num.index(num)
            break
    invoice_value = past_invoice_num[index_v:]
    invoice_num = int(invoice_value) + 1
    invoice_num_data.append(invoice_num)
    str_num = past_invoice_num[:index_v] + str(invoice_num)
    invoice_num_data.append(str_num)
    return invoice_num_data

# starts program
if __name__ == '__main__':
    Invoicemaker = True
    while Invoicemaker:
        filename = 'Invoice Excel Docs'
        directory = os.path.abspath(filename)
        invoice_data = {}
        # looks at all previous invoices and stores important customer data
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

        # make new invoice
        new_invoice = openpyxl.load_workbook('Invoice_template.xlsx')
        new_invoice_sheet = new_invoice.active
        # asks customer name to determine if customer is new or not
        name = input('Enter customer name: ')
        # if customer is not new, then it will use the customer data from the previous invoice
        if name in invoice_data:
            print('Customer name already exists')
            print('Making new invoice for customer from existing data')
            new_invoice_num = new_invoice_num()
            new_invoice_sheet['B1'] = new_invoice_num[1]
            new_invoice_sheet['B7'] = invoice_data[name]['customer_name']
            new_invoice_sheet['B8'] = invoice_data[name]['address']
            new_invoice_sheet['B9'] = invoice_data[name]['phone_number']
        # if customer is new, then it will ask for customer data
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

        # asks for invoice details
        new_invoice_sheet['C7'] = input('What is this invoice for? ')
        new_invoice_sheet['B13'] = input('Enter the details: ')
        new_invoice_sheet['C13'] = int(input('Enter the price: '))

        # saves new invoice separately and saves invoice template with updated invoice number
        new_invoice_template = new_invoice
        new_invoice_template.save('Invoice_template.xlsx')
        new_invoice.save(f'{name} #{new_invoice_num[0]}.xlsx')
        print(f'{new_invoice_num[0]} has been saved!')

        # determines if user wants to make another invoice
        continue_maker = input('Would you like to make another invoice? (y/n) ')
        # continues program if user wants to make another invoice
        if continue_maker == 'y':
            continue
        # exits program if user does not want to make another invoice
        else:
            print('Thank you for using the Invoice Maker!')
            Invoicemaker = False

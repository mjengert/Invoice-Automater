import openpyxl, os, shutil, tkinter
from tkinter import *
from tkinter import font as tkfont

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
        Invoicemaker = False

new_invoice = openpyxl.load_workbook('Invoice_template.xlsx')
new_invoice_sheet = new_invoice.active
name_list = []
invoice_info = []
InvoiceMaker = Tk()
InvoiceMaker.title("Invoice Maker")

# Set the size of the window
InvoiceMaker.geometry("550x350")


# Define a function for switching the frames
def change_to_input():
    data_input_page.pack(fill='both', expand=1)
    welcome_page.pack_forget()
    first_n = first_name.get()
    last_n = last_name.get()
    first_name.delete(0, END)
    last_name.delete(0, END)
    name_list.append(first_n)
    name_list.append(last_n)
    label3 = Label(data_input_page, text="Data Input!", font=font1)
    label3.pack(pady=10)
    label4 = Label(data_input_page, text="Input Customer's Information:", font=font2)
    label4.pack(pady=10)
    if name_list[0] + ' ' + name_list[1] in invoice_data:
        label5 = Label(data_input_page, text="Customer Data Was Found!", foreground="green", font=font2)
        label5.pack(pady=10)
        label6 = Label(data_input_page, text="Job Name: ", font=font2)
        label6.place(x=145, y=180)
        job_description = Entry(data_input_page, width=25)
        job_description.place(x=245, y=185)
        label7 = Label(data_input_page, text="Job Price: ", font=font2)
        label7.place(x=145, y=210)
        job_price = Entry(data_input_page, width=25)
        job_price.place(x=245, y=215)
        label8 = Label(data_input_page, text="Job Details: ", font=font2)
        label8.place(x=145, y=240)
        job_deet = Entry(data_input_page, width=25)
        job_deet.place(x=245, y=245)
        btn5 = Button(data_input_page, text="Submit", font=font2,
                      command=change_to_print(job_description, job_price, job_deet))
        btn5.pack(side=BOTTOM, pady=25)
    else:
        label9 = Label(data_input_page, text="Customer Data Was Not Found!", foreground="red", font=font2)
        label9.pack(pady=10)
        label10 = Label(data_input_page, text="Job Name: ", font=font2)
        label10.place(x=20, y=180)
        job_description = Entry(data_input_page, width=25)
        job_description.place(x=120, y=185)
        label11 = Label(data_input_page, text="Job Details: ", font=font2)
        label11.place(x=20, y=210)
        job_deet = Entry(data_input_page, width=25)
        job_deet.place(x=120, y=215)
        label12 = Label(data_input_page, text="Job price: ", font=font2)
        label12.place(x=145, y=250)
        job_price = Entry(data_input_page, width=25)
        job_price.place(x=245, y=255)
        cust_address = Label(data_input_page, text="Address: ", font=font2)
        cust_address.place(x=285, y=180)
        cust_address = Entry(data_input_page, width=25)
        cust_address.place(x=370, y=185)
        cust_phone = Label(data_input_page, text="Phone #: ", font=font2)
        cust_phone.place(x=285, y=210)
        cust_phone = Entry(data_input_page, width=25)
        cust_phone.place(x=370, y=215)
        btn6 = Button(data_input_page, text="Submit", font=font2,
                      command=change_to_print(job_description, job_price, job_deet, cust_address, cust_phone))
        btn6.pack(side=BOTTOM, pady=25)


def change_to_exit():
    exit_page.pack(fill='both', expand=1)
    print_page.pack_forget()


def change_to_welcome():
    welcome_page.pack(fill='both', expand=1)
    exit_page.pack_forget()


def change_to_print(job_description, job_price, job_deet, cust_address=0, cust_phone=0):
    print_page.pack(fill='both', expand=1)
    data_input_page.pack_forget()
    if name_list[0] + ' ' + name_list[1] in invoice_data:
        job_name = job_description.get()
        job_pricey = job_price.get()
        job_details = job_deet.get()
        invoice_info.append(job_name)
        invoice_info.append(job_details)
        invoice_info.append(job_pricey)
        PrintInvoice = Label(print_page, text="Would you like to print?", font=font1)
        PrintInvoice.pack(pady=10)
        btn7 = Radiobutton(print_page, text="Yes", font=font2, value=1)
        btn7.pack(pady=10)
    else:
        job_name = job_description.get()
        job_pricey = job_price.get()
        job_details = job_deet.get()
        address = cust_address.get()
        phone_number = cust_phone.get()
        invoice_info.append(job_name)
        invoice_info.append(job_details)
        invoice_info.append(job_pricey)
        invoice_info.append(address)
        invoice_info.append(phone_number)
        PrintInvoice = Label(print_page, text="Would you like to print?", font=font1)
        PrintInvoice.pack(pady=10)
        btn8 = Radiobutton(print_page, text="Yes", font=font2, value=1)
        btn8.pack(pady=10)


def save_invoice():
    name = name_list[0] + ' ' + name_list[1]
    print(invoice_info)
    # if customer is not new, then it will use the customer data from the previous invoice
    if name in invoice_data:
        new_invoice_num = new_invoice_numy()
        new_invoice_sheet['B1'] = new_invoice_num[1]
        new_invoice_sheet['B7'] = invoice_data[name]['customer_name']
        new_invoice_sheet['B8'] = invoice_data[name]['address']
        new_invoice_sheet['B9'] = invoice_data[name]['phone_number']
        new_invoice_sheet['C7'] = invoice_info[0]  # job name
        new_invoice_sheet['B13'] = invoice_info[1]  # job details
        new_invoice_sheet['C13'] = invoice_info[2]  # job price
    # if customer is new, then it will ask for customer data
    else:
        new_invoice_sheet['B7'] = name
        address = invoice_info[3]
        parts = address.split(", ")
        street = parts[0]
        city_state_zip = parts[1:]
        city, state, zip_code = city_state_zip[0], city_state_zip[1], city_state_zip[2]
        new_address = f"{street}\n{city}, {state.upper()} {zip_code}"
        new_invoice_sheet['B8'] = new_address
        new_invoice_num = new_invoice_numy()
        new_invoice_sheet['B1'] = new_invoice_num[1]
        new_invoice_sheet['B9'] = invoice_info[4]  # phone number
        invoice_data[name] = {'customer_name': name, 'address': new_address,
                              'phone_number': new_invoice_sheet['B9'].value}

        # asks for invoice details
        new_invoice_sheet['C7'] = invoice_info[0]  # job name
        new_invoice_sheet['B13'] = invoice_info[1]  # job details
        new_invoice_sheet['C13'] = invoice_info[2]  # job price

    # saves new invoice separately and saves invoice template with updated invoice number
    new_invoice_template = new_invoice
    new_invoice_template.save('Invoice_template.xlsx')
    new_invoice.save(f'{name} #{new_invoice_num[0]}.xlsx')


def new_invoice_numy():
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


# Create fonts for making difference in the frame
font1 = tkfont.Font(family='Helvetica', size='22', weight='bold')
font2 = tkfont.Font(family='Helvetica', size='12')

# Add a heading logo in the frames
# welcome frame widgets
welcome_page = Frame(InvoiceMaker)
data_input_page = Frame(InvoiceMaker)
print_page = Frame(InvoiceMaker)
exit_page = Frame(InvoiceMaker)
welcome_page.pack(fill='both', expand=1)
label1 = Label(welcome_page, text="Welcome to the Invoice Maker!", font=font1)
label1.pack(pady=20)
label2 = Label(welcome_page, text="Please enter the customer's first and last name:", foreground="blue", font=font2)
label2.pack(pady=20)
f_name = Label(welcome_page, text='First Name:', font=font2)
f_name.place(x=145, y=150)

first_name = Entry(welcome_page, width=25)
first_name.place(x=245, y=155)
l_name = Label(welcome_page, text='Last Name:', font=font2)
l_name.place(x=145, y=180)
last_name = Entry(welcome_page, width=25)
last_name.place(x=245, y=185)

first_n = first_name.get()
last_n = last_name.get()

btn1 = Button(welcome_page, text="Submit", font=font2, command=change_to_input)
btn1.pack(side=BOTTOM, pady=25)

# data input frame widgets
bt2 = Button(print_page, text="Submit", font=font2, command=change_to_exit)
bt2.pack(side=BOTTOM, pady=25)

# exit frame widgets
btn3 = Button(exit_page, text="Exit", font=font2, command=quit)
btn3.pack(pady=20)

btn4 = Button(exit_page, text="New Invoice", font=font2, command=change_to_welcome)
btn4.pack(pady=20)
InvoiceMaker.mainloop()

'''class InvoiceMaker(tkinter.Tk):
    def __init__(self, *args, **kwargs):
        tkinter.Tk.__init__(self, *args, **kwargs)
        self.title_font = tkfont.Font(family='Helvetica', size=20, weight='bold', slant='italic')
        self.smaller_font = tkfont.Font(family='Helvetica', size=12, weight='bold')
        main_frame = Frame(self)
        main_frame.pack(side='top', fill='both', expand=True)
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        self.first_name = None
        self.last_name = None

        self.frames = {}
        for F in (WelcomePage, InvoiceData):
            page_name = F.__name__
            frame = F(parent=main_frame, controller=self)
            self.frames[page_name] = frame
            frame.grid(row=0, column=0, sticky='nsew')

        self.show_screen('WelcomePage')

    def show_screen(self, page_name):
        frame = self.frames[page_name]
        frame.tkraise()


class WelcomePage(tkinter.Frame):
    def __init__(self, parent, controller):
        tkinter.Frame.__init__(self, parent)
        self.controller = controller
        label = tkinter.Label(self, text='Welcome to the Invoice Maker!', font=controller.title_font)
        label.pack(side='top', fill='x', pady=20)
        instruct = tkinter.Label(self, text="Please enter the customer's first and last name:",
                                 font=controller.smaller_font)
        instruct.pack(side='top', fill='x', pady=10)
        f_name = Label(self, text='First Name:')
        l_name = Label(self, text='Last Name:')
        f_name.place(x=15, y=130)
        l_name.place(x=15, y=155)
        self.first_name_entry = Entry(self, bd=5)
        self.last_name_entry = Entry(self, bd=5)
        self.first_name_entry.place(x=95, y=130)
        self.last_name_entry.place(x=95, y=155)
        button1 = Label(self, text='')
        button1.pack(side='right', pady=20)
        button2 = tkinter.Button(self, text='Submit', command=self.submit_name)
        button2.place(x=285, y=143.5)

    def submit_name(self):
        InvoiceData.first_name = self.first_name_entry.get()
        InvoiceData.last_name = self.last_name_entry.get()
        self.controller.show_screen('InvoiceData')


class InvoiceData(tkinter.Frame):
    def __init__(self, parent, controller):
        tkinter.Frame.__init__(self, parent)
        self.controller = controller
        self.first_name = self.controller.first_name
        self.last_name = self.controller.last_name
        truth = True
        while truth:
            print("First Name: ", self.first_name)
            print("Last Name: ", self.last_name)
        label = tkinter.Label(self, text='Enter Customer Data', font=controller.title_font)
        label.pack(side='top', fill='x', pady=10)
        button1 = tkinter.Button(self, text='Submit', command=quit)
        button1.pack()

    def __str__(self):
        return self.controller.first_name + ' ' + self.controller.last_name


if __name__ == "__main__":
    app = InvoiceMaker()
    app.mainloop()'''

import openpyxl, os, shutil, tkinter
from tkinter import *
from tkinter import font as tkfont


class InvoiceMaker(tkinter.Tk):
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
    app.mainloop()


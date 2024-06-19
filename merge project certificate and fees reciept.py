#!/usr/bin/env python
# coding: utf-8

# In[27]:


import tkinter
from tkinter import ttk
from tkinter import *
from docxtpl import DocxTemplate
from tkcalendar import Calendar
from datetime import datetime
from docx2pdf import convert

window = tkinter.Tk()
window.title("Welcome to edS/ash")

frame = tkinter.Frame(window)
frame.grid(padx = 20,pady = 20)

def exit_app():
    window.destroy()

def generat_certificate():
    
    doc = DocxTemplate("edS Certificate.docx")
    name = sname_entry.get()
    application_no = application_entry.get()
    cname = combo_box.get()
    now = datetime.now()
    current_date = now.date()
    formatted_date = current_date.strftime('%d %B %Y')
    
    doc.render({"name" :name,
                "application_no" : application_no,
            "course_name" :cname,
            "date" : formatted_date})
    
    exit_app()

    doc_name = name  + ".docx"
    doc.save(doc_name)
    # pdf_name = name  + ".docx"
    # convert(pdf_name)
    
    print("Certificate Generate Successful ")
    
    exit_app()
    
def certificate():
    window1 = tkinter.Tk()
    window1.title("GC")

    frame1 = tkinter.Frame(window1)
    frame1.grid(padx = 20,pady = 20)
    
    sname_label = tkinter.Label(frame1, text = "Name")
    sname_label.grid(row = 0,column = 0,padx = 20, pady=5,sticky ="w")  # to fix the position os frame      ((((label is used for label means to display))))
    sname_entry = tkinter.Entry(frame1)
    sname_entry.grid(row = 0,column = 1,padx = 20, pady=5)  # to fix the position os frame      ((((entry is used to take input))))

    application_lebel = tkinter.Label(frame1, text = "Application no.")
    application_lebel.grid(row =1 ,column = 0,padx = 20,pady = 5,sticky = "w")
    application_entry = tkinter.Entry(frame1)
    application_entry.grid(row = 1, column = 1, padx = 50,pady = 5)

    course_label = tkinter.Label(frame1, text = "Course Name")
    course_label.grid(row = 2,column = 0,padx = 20, pady=5,sticky ="w")  # to fix the position os frame
    combo_box = ttk.Combobox(frame1, values=["Basic and Advance C programming","Basic and Advance C++ programming","Basic and Advance Python programming","Basic and Advance SQL programming","Data Science","Data Analysis","Machine Learning","Generative AI",])
    combo_box.grid(row = 2,column = 1,padx = 20, pady=5)

    generate_receipt_button = tkinter.Button(frame1 , text = "generate_certificate", command = generat_certificate)      # command is used to call a function (function name is add item) # this is used to create button 
    generate_receipt_button.grid(row=3, column=0, columnspan = 2, sticky ="news",padx = 20, pady=5) 

    # window1.mainloop()

##################     Fees
def Fees_Receipt():
    window1 = tkinter.Tk()
    window1.title("Fees_Receipt")

    frame1 = tkinter.Frame(window1) # frame is a container inside the window container (frame store other information) # bracket me window is leye lekha because it is parent window
    frame1.pack(padx=40, pady=20)               # pack to use for grid or we can se to fix length width   # frame window ke under h


    def generate_receipt():
        doc = DocxTemplate("Fees_Receipt_template.docx")
        name = sname_entry.get()
        #gender = gen_mode.get()

        gender_options = {1: "Male", 2: "Female"}
        gender = gender_options[v.get()]

        application_no = application_entry.get()

        cname = combo_box.get()
        #cname = selected_item

        mode_options = {1: "Cash", 2: "UPI"}
        pmode = mode_options[m.get()]

        amount = fees_entry.get()

        ddate = date_combobox.get()
        #rno = application_no+"/"+sname

        doc.render({"name" :name,
                "gen" : gender,
                "application_no" :application_no,
                "course" :cname,
                "mode" :pmode,
                "amount" :amount + "rs",
                "date" : ddate})

        doc_name = name  + ".docx"
        doc.save(doc_name)

        # pdf_name = name  + ".docx"
        # convert(pdf_name)
        exit_app()


    sname_label = tkinter.Label(frame1, text = "Name")
    sname_label.grid(row = 0,column = 0,padx = 50, pady=5)  # to fix the position os frame      ((((label is used for label means to display))))
    sname_entry = tkinter.Entry(frame1)
    sname_entry.grid(row = 0,column = 1,padx = 50, pady=5)  # to fix the position os frame      ((((entry is used to take input))))

    gen_lebel = tkinter.Label(frame1, text = "GENDER")
    gen_lebel.grid(row = 1,column = 0,padx = 50, pady=5)
    v = IntVar()
    gen_mode = tkinter.Radiobutton(frame1, text='Male', variable=v, value=1).grid(row=1,column = 1, sticky=W,padx = 50, pady=5)
    gen_mode = tkinter.Radiobutton(frame1, text='Female', variable=v, value=2).grid(row=1,column =2, sticky=W,padx = 50, pady=5)

    application_label = tkinter.Label(frame1, text = "Application no.")
    application_label.grid(row = 2,column = 0,padx = 50, pady=5)  # to fix the position os frame
    application_entry = tkinter.Entry(frame1)
    application_entry.grid(row = 2,column = 1,padx = 50, pady=5)  # to fix the position os frame      ((((entry is used to take input))))

    course_label = tkinter.Label(frame1, text = "Course Name")
    course_label.grid(row = 3,column = 0,padx = 50, pady=5)  # to fix the position os frame

    # Create a Combobox widget
    combo_box = ttk.Combobox(frame1, values=["Data Science", "Machine Learning", "Python"])
    combo_box.grid(row = 3,column = 1,padx = 50, pady=5)

    payment_lebel = tkinter.Label(frame1, text = "Payment Mode")
    payment_lebel.grid(row = 4,column = 0,padx = 50, pady=5)
    m = IntVar()
    payment = tkinter.Radiobutton(frame1, text='Cash', variable=m, value=1).grid(row=4,column = 1, sticky=W,padx = 50, pady=5)
    payment = tkinter.Radiobutton(frame1, text='UPI', variable=m, value=2).grid(row=4,column =2, sticky=W,padx = 50, pady=5)

    fees_label = tkinter.Label(frame1, text = "Fees Amount")
    fees_label.grid(row = 5,column = 0,padx = 50, pady=5)  # to fix the position os frame

    fees_entry = tkinter.Entry(frame1)
    fees_entry.grid(row = 5,column = 1,padx = 50, pady=5)

    date_label = tkinter.Label(frame1, text = "Date")
    date_label.grid(row = 6,column = 0,padx = 50, pady=5)  # to fix the position os frame


    # Function to be called when combobox is clicked
    def show_calendar():
        root = Tk()
        root.title("Select date")

        # Get the current date
        current_date = datetime.now()
        current_year = current_date.year
        current_month = current_date.month
        current_day = current_date.day

        # Add a Calendar widget to the popup window
        cal = Calendar(root, selectmode='day', year=current_year, month=current_month, day=current_day)
        cal.grid(row = 0,column = 0,padx = 5, pady=5)

        # Function to get the selected date from the calendar and set it in the combobox
        def get_date():
            selected_date = cal.get_date()
            date_combobox.set(selected_date)
            root.destroy()  # Close the popup window

        # Add a Button to the popup window to confirm the date selection
        Button = tkinter.Button(root, text="Select Date", command=get_date)
        Button.grid(row = 1,column = 0,padx = 5, pady=5)

    # Create a combobox widget
    date_combobox = ttk.Combobox(frame1)
    date_combobox.grid(row = 6,column = 1,padx = 5, pady=5)

    # Bind the combobox to show the calendar when clicked
    date_combobox.bind('<Button-1>', lambda event: show_calendar())

    generate_receipt_button = tkinter.Button(frame1 , text = "generate_receipt", command = generate_receipt)      # command is used to call a function (function name is add item) # this is used to create button 
    generate_receipt_button.grid(row=7, column=0, columnspan = 3, sticky ="news",padx = 50, pady=5) 


    # window1.mainloop() 
    
###################    

def submit():
    chh_options = {1: "Cretificate", 2: "Fees Reciept"}
    ch = chh_options[v.get()]
    if ch == "Cretificate":
        certificate()
    else :
        Fees_Receipt()
    
    
    
    
    
    
Ch_lebel = tkinter.Label(frame, text = "Choice")
Ch_lebel.grid(row = 0,column = 0,padx = 50, pady=5)
v = IntVar()
Ch_mode = tkinter.Radiobutton(frame, text='Cretificate', variable=v, value=1).grid(row=0,column = 1, sticky=W,padx = 50, pady=5)
Ch_mode = tkinter.Radiobutton(frame, text='Fees Reciept', variable=v, value=2).grid(row=0,column =2, sticky=W,padx = 50, pady=5)



submit_button = tkinter.Button(frame , text = "Submit", command = submit)      # command is used to call a function (function name is add item) # this is used to create button 
submit_button.grid(row=1, column=0, columnspan = 3, sticky ="news",padx = 50, pady=5) 

    
window.mainloop()


# In[25]:


######################         Single certificate               #####################
import tkinter
from tkinter import ttk
from tkinter import *
from docxtpl import DocxTemplate
from tkcalendar import Calendar
from datetime import datetime
from docx2pdf import convert

window = tkinter.Tk()
window.title("Generate edSlash Certificate")
#window.geometry("400x400")

frame = tkinter.Frame(window) # frame is a container inside the window container (frame store other information) # bracket me window is leye lekha because it is parent window
frame.pack(padx=60, pady=40)               # pack to use for grid or we can se to fix length width   # frame window ke under h

def exit_app():
    window.destroy()

def generat_certificate():
    
    doc = DocxTemplate("edS Certificate.docx")
    name = sname_entry.get()
    application_no = application_entry.get()
    cname = combo_box.get()
    now = datetime.now()
    current_date = now.date()
    formatted_date = current_date.strftime('%d %B %Y')
    
    doc.render({"name" :name,
                "application_no" : application_no,
            "course_name" :cname,
            "date" : formatted_date})
    
    exit_app()

    doc_name = name  + ".docx"
    doc.save(doc_name)
    pdf_name = name  + ".docx"
    convert(pdf_name)
    
    print("Certificate Generate Successful ")

sname_label = tkinter.Label(frame, text = "Name")
sname_label.grid(row = 0,column = 0,padx = 20, pady=5,sticky ="w")  # to fix the position os frame      ((((label is used for label means to display))))
sname_entry = tkinter.Entry(frame)
sname_entry.grid(row = 0,column = 1,padx = 20, pady=5)  # to fix the position os frame      ((((entry is used to take input))))

application_lebel = tkinter.Label(frame, text = "Application no.")
application_lebel.grid(row =1 ,column = 0,padx = 20,pady = 5,sticky = "w")
application_entry = tkinter.Entry(frame)
application_entry.grid(row = 1, column = 1, padx = 50,pady = 5)

course_label = tkinter.Label(frame, text = "Course Name")
course_label.grid(row = 2,column = 0,padx = 20, pady=5,sticky ="w")  # to fix the position os frame
combo_box = ttk.Combobox(frame, values=["Basic and Advance C programming","Basic and Advance C++ programming","Basic and Advance Python programming","Basic and Advance SQL programming","Data Science","Data Analysis","Machine Learning","Generative AI",])
combo_box.grid(row = 2,column = 1,padx = 20, pady=5)

generate_receipt_button = tkinter.Button(frame , text = "generate_certificate", command = generat_certificate)      # command is used to call a function (function name is add item) # this is used to create button 
generate_receipt_button.grid(row=3, column=0, columnspan = 2, sticky ="news",padx = 20, pady=5) 

window.mainloop()


# In[ ]:


######################         multiple certificate generate using excel                 #####################
import tkinter
import numpy as np
import pandas as pd
from tkinter import ttk
from tkinter import *
from tkinter import filedialog  # Import filedialog
from docxtpl import DocxTemplate
from tkcalendar import Calendar
from datetime import datetime
from docx2pdf import convert

def generat_certificate(names,application_no,cname):
    
    doc = DocxTemplate("edS Certificate.docx")
    
    now = datetime.now()
    current_date = now.date()
    formatted_date = current_date.strftime('%d %B %Y')
    
    doc.render({"name" :names,
                "application_no" : application_no,
            "course_name" :cname,
            "date" : formatted_date})

    doc_name = names  + ".docx"
    doc.save(doc_name)
    pdf_name = names  + ".docx"
    convert(pdf_name)
    
    print("Certificate Generate Successful ")





def browseFiles():
    global df
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select a File",
                                          filetypes=(("Excel files", "*.xlsx"),
                                                     ("CSV files", "*.csv")))
    label_file_explorer.configure(text="File Opened: " + filename)
    if filename.endswith('.csv'):
        df = pd.read_csv(filename)
    elif filename.endswith('.xlsx'):
        df = pd.read_excel(filename)

def exit_app():
    window.destroy()

window = tkinter.Tk()
window.title('File Explorer')
#window.geometry("500x500")
window.config(background="white")

label_file_explorer = Label(window, text="File Explorer using Tkinter", width=50, height=4, fg="blue")
label_file_explorer.grid(column=0, row=1)

button_explore = Button(window, text="Browse Files", command=browseFiles)
button_explore.grid(column=0, row=2,padx = 50,pady = 10)

button_exit = Button(window, text="Exit", command=exit_app)
button_exit.grid(column=0, row=3,padx = 50,pady = 10)

window.mainloop()

# Ensure 'df' is defined before printing columns
try:
    print(df.columns)
except NameError:
    print("No file loaded or 'df' is not defined.")

for i in range(len(df)):
    application_no = df["Application_no"].iloc[i]
    names = df["Name"].iloc[i]
    cname = df["Course_name"].iloc[i]
    generat_certificate(names,application_no,cname)


# In[1]:


######################            Fees reciept generate              #####################

import tkinter
import numpy as np
import pandas as pd
from tkinter import ttk
from tkinter import *
from tkinter import filedialog  # Import filedialog
from docxtpl import DocxTemplate
from tkcalendar import Calendar
from datetime import datetime
from docx2pdf import convert

window = tkinter.Tk()
window.title("Fees_Receipt")

frame = tkinter.Frame(window) # frame is a container inside the window container (frame store other information) # bracket me window is leye lekha because it is parent window
frame.pack(padx=40, pady=20)               # pack to use for grid or we can se to fix length width   # frame window ke under h


def generate_receipt():
    doc = DocxTemplate("Fees_Receipt_template.docx")
    name = sname_entry.get()
    #gender = gen_mode.get()
    
    gender_options = {1: "Male", 2: "Female"}
    gender = gender_options[v.get()]
    
    application_no = application_entry.get()
    
    cname = combo_box.get()
    #cname = selected_item
    
    mode_options = {1: "Cash", 2: "UPI"}
    pmode = mode_options[m.get()]
    
    amount = fees_entry.get()
    
    ddate = date_combobox.get()
    #rno = application_no+"/"+sname
    
    doc.render({"name" :name,
            "gen" : gender,
            "application_no" :application_no,
            "course" :cname,
            "mode" :pmode,
            "amount" :amount + "rs",
            "date" : ddate})

    doc_name = name  + ".docx"
    doc.save(doc_name)
    
    pdf_name = name  + ".docx"
    convert(pdf_name)


sname_label = tkinter.Label(frame, text = "Name")
sname_label.grid(row = 0,column = 0,padx = 50, pady=5)  # to fix the position os frame      ((((label is used for label means to display))))
sname_entry = tkinter.Entry(frame)
sname_entry.grid(row = 0,column = 1,padx = 50, pady=5)  # to fix the position os frame      ((((entry is used to take input))))

gen_lebel = tkinter.Label(frame, text = "GENDER")
gen_lebel.grid(row = 1,column = 0,padx = 50, pady=5)
v = IntVar()
gen_mode = tkinter.Radiobutton(frame, text='Male', variable=v, value=1).grid(row=1,column = 1, sticky=W,padx = 50, pady=5)
gen_mode = tkinter.Radiobutton(frame, text='Female', variable=v, value=2).grid(row=1,column =2, sticky=W,padx = 50, pady=5)

application_label = tkinter.Label(frame, text = "Application no.")
application_label.grid(row = 2,column = 0,padx = 50, pady=5)  # to fix the position os frame
application_entry = tkinter.Entry(frame)
application_entry.grid(row = 2,column = 1,padx = 50, pady=5)  # to fix the position os frame      ((((entry is used to take input))))

course_label = tkinter.Label(frame, text = "Course Name")
course_label.grid(row = 3,column = 0,padx = 50, pady=5)  # to fix the position os frame

# course_entry = tkinter.Entry(frame)
# course_entry.grid(row = 3,column = 1)


# def on_select(event):
#     selected_item = combo_box.get()
#     label.config(text="Selected Item: " + selected_item)

# Create a label
# label = tkinter.Label(frame, text="Selected Item: ")
# label.grid(pady=10)

# Create a Combobox widget
combo_box = ttk.Combobox(frame, values=["Data Science", "Machine Learning", "Python"])
combo_box.grid(row = 3,column = 1,padx = 50, pady=5)

# Set default value
#combo_box.set("Option 1")

# Bind event to selection
#combo_box.bind("<<ComboboxSelected>>", on_select)

payment_lebel = tkinter.Label(frame, text = "Payment Mode")
payment_lebel.grid(row = 4,column = 0,padx = 50, pady=5)
m = IntVar()
payment = tkinter.Radiobutton(frame, text='Cash', variable=m, value=1).grid(row=4,column = 1, sticky=W,padx = 50, pady=5)
payment = tkinter.Radiobutton(frame, text='UPI', variable=m, value=2).grid(row=4,column =2, sticky=W,padx = 50, pady=5)

fees_label = tkinter.Label(frame, text = "Fees Amount")
fees_label.grid(row = 5,column = 0,padx = 50, pady=5)  # to fix the position os frame

fees_entry = tkinter.Entry(frame)
fees_entry.grid(row = 5,column = 1,padx = 50, pady=5)



date_label = tkinter.Label(frame, text = "Date")
date_label.grid(row = 6,column = 0,padx = 50, pady=5)  # to fix the position os frame


# Function to be called when combobox is clicked
def show_calendar():
    root = Tk()
    root.title("Select date")

# Set the geometry of the window
    #root.geometry("400x400")

    # Get the current date
    current_date = datetime.now()
    current_year = current_date.year
    current_month = current_date.month
    current_day = current_date.day

    # Add a Calendar widget to the popup window
    cal = Calendar(root, selectmode='day', year=current_year, month=current_month, day=current_day)
    cal.grid(row = 0,column = 0,padx = 5, pady=5)

    # Function to get the selected date from the calendar and set it in the combobox
    def get_date():
        selected_date = cal.get_date()
        date_combobox.set(selected_date)
        root.destroy()  # Close the popup window

    # Add a Button to the popup window to confirm the date selection
    Button = tkinter.Button(root, text="Select Date", command=get_date)
    Button.grid(row = 1,column = 0,padx = 5, pady=5)
    
# Create a combobox widget
date_combobox = ttk.Combobox(frame)
date_combobox.grid(row = 6,column = 1,padx = 5, pady=5)

# Bind the combobox to show the calendar when clicked
date_combobox.bind('<Button-1>', lambda event: show_calendar())
    
generate_receipt_button = tkinter.Button(frame , text = "generate_receipt", command = generate_receipt)      # command is used to call a function (function name is add item) # this is used to create button 
generate_receipt_button.grid(row=7, column=0, columnspan = 3, sticky ="news",padx = 50, pady=5) 


window.mainloop()


# In[2]:


import tkinter
import numpy as np
import pandas as pd
from tkinter import ttk
from tkinter import *
from tkinter import filedialog
from docxtpl import DocxTemplate
from tkcalendar import Calendar
from datetime import datetime
from docx2pdf import convert

window = tkinter.Tk()
window.title("Welcome to edS/ash")

frame = tkinter.Frame(window)
frame.grid(padx=20, pady=20)

def exit_app():
    window.destroy()

def generat_certificate():
    doc = DocxTemplate("edS Certificate.docx")
    name = sname_entry.get()
    application_no = application_entry.get()
    cname = course_box.get()
    now = datetime.now()
    current_date = now.date()
    formatted_date = current_date.strftime('%d %B %Y')

    doc.render({
        "name": name,
        "application_no": application_no,
        "course_name": cname,
        "date": formatted_date})

    doc_name = name + " certificate.docx"
    doc.save(doc_name) # to save word file
    # Uncomment the lines below to enable PDF conversion
    pdf_name = name + " certificate.docx"
    convert(pdf_name)

    print(f"{name} Certificate Generated Successfully")
    exit_app()

def certificate():
    global sname_entry, application_entry, combo_box
    window1 = tkinter.Tk()
    window1.title("GC")

    frame1 = tkinter.Frame(window1)
    frame1.grid(padx=20, pady=20)

    sname_label = tkinter.Label(frame1, text="Name")
    sname_label.grid(row=0, column=0, padx=20, pady=5, sticky="w")
    sname_entry = tkinter.Entry(frame1)
    sname_entry.grid(row=0, column=1, padx=20, pady=5)

    application_label = tkinter.Label(frame1, text="Application no.")
    application_label.grid(row=1, column=0, padx=20, pady=5, sticky="w")
    application_entry = tkinter.Entry(frame1)
    application_entry.grid(row=1, column=1, padx=20, pady=5)

    course_label = tkinter.Label(frame1, text="Course Name")
    course_label.grid(row=2, column=0, padx=20, pady=5, sticky="w")
    course_box = ttk.Combobox(frame1, values=["Basic and Advance C programming", "Basic and Advance C++ programming",
        "Basic and Advance Python programming", "Basic and Advance SQL programming",
        "Data Science", "Data Analysis", "Machine Learning", "Generative AI"])
    course_box.grid(row=2, column=1, padx=20, pady=5)

    generate_receipt_button = tkinter.Button(frame1, text="Generate Certificate", command=generat_certificate)
    generate_receipt_button.grid(row=3, column=0, columnspan=2, sticky="news", padx=20, pady=5)

def Fees_Receipt():
    global sname_entry, application_entry, combo_box, fees_entry, date_combobox
    window1 = tkinter.Tk()
    window1.title("Fees Receipt")

    frame1 = tkinter.Frame(window1)
    frame1.pack(padx=40, pady=20)

    def generate_receipt():
        doc = DocxTemplate("Fees_Receipt_template.docx")
        name = sname_entry.get()
        gen = gen_mode.get()
        application_no = application_entry.get()
        cname = course_box.get()
        pmode = payment_mode.get()
        amount = fees_entry.get()
        ddate = date_combobox.get()

        doc.render({
            "name": name,
            "gen": gender,
            "application_no": application_no,
            "course": cname,
            "mode": pmode,
            "amount": "₹" + amount + "/-",
            "date": ddate})

        doc_name = name + " Fees_Receipt.docx"
        doc.save(doc_name)
        # Uncomment the lines below to enable PDF conversion
        pdf_name = name + " Fees_Receipt.docx"
        convert(pdf_name)
        
        print(f"{name} Fees_Receipt Generated Successfully")
        
        exit_app()

    sname_label = tkinter.Label(frame1, text="Name")
    sname_label.grid(row=0, column=0, padx=50, pady=5)
    sname_entry = tkinter.Entry(frame1)
    sname_entry.grid(row=0, column=1, padx=50, pady=5)
    
    gen_label = tkinter.Label(frame1, text="Payment Mode")
    gen_label.grid(row=1, column=0, padx=50, pady=5)
    gen_mode = ttk.Combobox(frame1, values=["Cash", "UPI"])
    gen_mode.grid(row=1, column=1, padx=50, pady=5)

    application_label = tkinter.Label(frame1, text="Application no.")
    application_label.grid(row=2, column=0, padx=50, pady=5)
    application_entry = tkinter.Entry(frame1)
    application_entry.grid(row=2, column=1, padx=50, pady=5)

    course_label = tkinter.Label(frame1, text="Course Name")
    course_label.grid(row=3, column=0, padx=50, pady=5)
    course_box = ttk.Combobox(frame1, values=["Data Science", "Machine Learning", "Python"])
    course_box.grid(row=3, column=1, padx=50, pady=5)
    
    payment_label = tkinter.Label(frame1, text="Payment Mode")
    payment_label.grid(row=4, column=0, padx=50, pady=5)
    payment_mode = ttk.Combobox(frame1, values=["Cash", "UPI"])
    payment_mode.grid(row=4, column=1, padx=50, pady=5)

    fees_label = tkinter.Label(frame1, text="Fees Amount")
    fees_label.grid(row=5, column=0, padx=50, pady=5)
    fees_entry = tkinter.Entry(frame1)
    fees_entry.grid(row=5, column=1, padx=50, pady=5)

    date_label = tkinter.Label(frame1, text="Date")
    date_label.grid(row=6, column=0, padx=50, pady=5)

    def show_calendar():
        root = Tk()
        root.title("Select Date")
        current_date = datetime.now()
        cal = Calendar(root, selectmode='day', year=current_date.year, month=current_date.month, day=current_date.day)
        cal.grid(row=0, column=0, padx=5, pady=5)

        def get_date():
            selected_date = cal.get_date()
            date_combobox.set(selected_date)
            root.destroy()

        select_date_button = tkinter.Button(root, text="Select Date", command=get_date)
        select_date_button.grid(row=1, column=0, padx=5, pady=5)

    date_combobox = ttk.Combobox(frame1)
    date_combobox.grid(row=6, column=1, padx=5, pady=5)
    date_combobox.bind('<Button-1>', lambda event: show_calendar())

    generate_receipt_button = tkinter.Button(frame1, text="Generate Receipt", command=generate_receipt)
    generate_receipt_button.grid(row=7, column=0, columnspan=3, sticky="news", padx=50, pady=5)

def generat_multiple_certificate():
    def generat_certificate(names,application_no,cname):

        doc = DocxTemplate("edS Certificate.docx")

        now = datetime.now()
        current_date = now.date()
        formatted_date = current_date.strftime('%d %B %Y')

        doc.render({"name" :names,
                    "application_no" : application_no,
                "course_name" :cname,
                "date" : formatted_date})

        doc_name = names  + ".docx"
        doc.save(doc_name)
        pdf_name = names  + ".docx"
        convert(pdf_name)

        print(f"{names} Certificate Generate Successful ")

    def browseFiles():
        global df
        filename = filedialog.askopenfilename(initialdir="/",
                                              title="Select a File",
                                              filetypes=(("Excel files", "*.xlsx"),
                                                         ("CSV files", "*.csv")))
        label_file_explorer.configure(text="File Opened: " + filename)
        if filename.endswith('.csv'):
            df = pd.read_csv(filename)
        elif filename.endswith('.xlsx'):
            df = pd.read_excel(filename)

    def exit_app():
        window1.destroy()
          
    window1 = tkinter.Tk()
    window1.title("Multiple Certificate")

    frame1 = tkinter.Frame(window1)
    frame1.pack(padx=40, pady=20)

    label_file_explorer = Label(frame1, text="File Explorer using Tkinter", width=50, height=4, fg="blue")
    label_file_explorer.grid(column=0, row=1)

    button_explore = Button(frame1, text="Browse Files", command=browseFiles)
    button_explore.grid(column=0, row=2,padx = 50,pady = 10)

    button_exit = Button(frame1, text="Exit", command=exit_app)
    button_exit.grid(column=0, row=3,padx = 50,pady = 10)

    try:
        print(df.columns)
    except NameError:
        print("No file loaded or 'df' is not defined.")

    for i in range(len(df)):
        application_no = df["Application_no"].iloc[i]
        names = df["Name"].iloc[i]
        cname = df["Course_name"].iloc[i]
        generat_certificate(names,application_no,cname)

def submit():
    chh_options = {1: "Certificate",2: 'Multiple certificate',3: "Fees receipt"}
    ch = chh_options[v.get()]
    if ch == "Certificate":
        certificate()
    elif ch == 'Multiple certificate':
        generat_multiple_certificate()
    elif ch == "Fees receipt":
        Fees_Receipt()

ch_label = tkinter.Label(frame, text="Choice: ")
ch_label.grid(row=0, column=0, padx=20, pady=5)
v = IntVar()
ch_mode_certificate = tkinter.Radiobutton(frame, text='Certificate', variable=v, value=1)
ch_mode_certificate.grid(row=0, column=1, sticky=W, padx=10, pady=5)

ch_mode_multiple_certificate = tkinter.Radiobutton(frame, text='Multiple Certificate', variable=v, value=2)
ch_mode_multiple_certificate.grid(row=1, column=1, sticky=W, padx=10, pady=5)

ch_mode_fees_receipt = tkinter.Radiobutton(frame, text='Fees Receipt', variable=v, value=3)
ch_mode_fees_receipt.grid(row=2, column=1, sticky=W, padx=10, pady=5)

submit_button = tkinter.Button(frame, text="Submit", command=submit)
submit_button.grid(row=3, column=1, sticky="news", padx=50, pady=5)

window.mainloop()


# In[ ]:





# In[ ]:





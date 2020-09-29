import cv2
import pytesseract
import pandas as pd
import tkinter as tk
from tkinter import filedialog, Text, Label, OptionMenu
import pathlib
from csv2vcard import csv2vcard
import vobject
from csv import writer
from sys import stdout
import os



OPTIONS = [
    "eng",
    "heb"
]
arr = []
all_linesA = []

out = open("cards.csv", "w+")
all_linesA.append("last_name")
all_linesA.append("first_name")
all_linesA.append("org")
all_linesA.append("title")
all_linesA.append("phone")
all_linesA.append("email")
all_linesA.append("website")
all_linesA.append("street")
all_linesA.append("city")
all_linesA.append("p_code")
all_linesA.append("country")
out.writelines(all_linesA)


def print_file():
    filename = filedialog.askopenfilename(initialdir="/Users/jacobsongal/Documents/General/Programs/Pycharm/easyContects/resources", title="Selcet File")
    ftype = pathlib.Path(filename).suffix

    if ftype.__eq__('.jpg') or ftype.__eq__('.png') or ftype.__eq__('.jpeg'):
        img = cv2.imread(filename)
        names = pytesseract.pytesseract.image_to_string(img, config='digits')
        temp = pytesseract.pytesseract.image_to_string(img, lang=variable.get())
        print(names)
        print(temp)
        for index in names:
            label = app.Label(app, text=names[index], bg="gray")
            label.place(relwidth=0.8, relheight=0.1, relx=0, rely=0.1 + index / 10 + 0.1)

    elif ftype.__eq__('.xls') or ftype.__eq__('.xl'):
        datafile = pd.read_excel(filename, header=0)
        arr = datafile.to_dict('index')
        for index in arr:
            label = tk.Label(app, text=arr[index], bg="white")
            label.place(relwidth=0.7, relheight=0.05, relx=0, rely=0.15+index/10+0.1)
            indexbut = tk.Button(app, text="Add Contact", padx=10, pady=5, fg="black", bg="white", command=vcard_factory)
            indexbut.place(relwidth=0.15, relheight=0.05, relx=0.7, rely=0.15+index/10+0.1)

            print(arr[index])

            card = vobject.vCard()
            given = arr[index]['שם פרטי']
            family = arr[index]['שם משפחה ']
            phone = arr[index]['נייד']
            card.add('fn').value = given+' '+family
            card.add('phone').value = phone
            print(card.serialize())

            out = open("cards.csv", "a+")
            all_lines = []
            all_lines.append(family)
            all_lines.append(",")
            all_lines.append(given)
            all_lines.append(",")
            all_lines.append("army")
            all_lines.append(",")
            all_lines.append("title")
            all_lines.append(",")
            all_lines.append("email")
            all_lines.append(",")
            all_linesA.append("website")
            all_lines.append(",")
            all_linesA.append("street")
            all_lines.append(",")
            all_linesA.append("city")
            all_lines.append(",")
            all_linesA.append("p_code")
            all_lines.append(",")
            all_linesA.append("country")

            out.writelines(all_lines)

            csv2vcard.csv2vcard("cards.csv", ",")




def vcard_factory():
        item = arr[0]
        v = vobject.vCard()
        v.add('fn').value = item.pop(1)
        v.add('phone').value = item.pop(2)
        # v.add('email').value = 'jeffrey@example.com'
        print(v.serialize())


def add_all():
    for index in arr:
        item=arr[index]
        v = vobject.vCard()
        v.add('fn').value = item.pop(1)
        v.add('phone').value = item.pop(2)
        # v.add('email').value = 'jeffrey@example.com'
        print(v.serialize())


app = tk.Tk()
app.geometry('700x700')
head = Label(app, text="Easy Contects")
head.pack()

canvas = tk.Canvas(app, height=700, width=700, bg="white")
canvas.pack()

#frame = tk.Frame(app, bg="white")
#frame.place(relwidth=0.8, relheight=0.8, relx=0.1, rely=0.1)

variable = tk.StringVar(app)
variable.set(OPTIONS[0])

language = tk.OptionMenu(app, variable, *OPTIONS)
language.config(width=90, font=('Helvetica', 12))
language.place(relwidth=0.1, relheight=0.04, relx=0.1, rely=0.2)

openfile = tk.Button(app, text="Open File", padx=10, pady=5, fg="black", bg="white", command=print_file)
openfile.place(relwidth=0.1, relheight=0.1, relx=0.1, rely=0.1)

addContact = tk.Button(app, text="Add all Contacts", padx=10, pady=5, fg="black", bg="white", command=add_all)
addContact.place(relwidth=0.15, relheight=0.1, relx=0.2, rely=0.1)

app.mainloop()


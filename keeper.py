from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from os.path import join, dirname, abspath
import xlrd


class GUI:
    window = Tk()
    window.title("AttendanceKeeper V1.0")

    mainframe = ttk.Frame(window, padding="3 3 12 12")
    mainframe.grid(column=3, row=7, sticky=(N, W, E, S))

    excelPath = ""

    def dosya():
        global window, excelPath
        path = filedialog.askopenfilename(title="Select file")
        excelPath = path

    lbl1 = Label(mainframe, text="Attendance Keeper V1.0", font=("Arial Bold", 25))
    lbl2 = Label(mainframe, text="Select students list Excel file", font=("Arial Bold", 15))
    lbl3 = Label(mainframe, text="Select students:", font=("Arial Bold", 15))
    lbl4 = Label(mainframe, text="Section:", font=("Arial Bold", 15))
    lbl5 = Label(mainframe, text="Attended Studets:", font=("Arial Bold", 15))

    button = ttk.Button(master=mainframe, text="Hello", command=dosya)

    lbl1.grid(column=2, row=1, sticky="N",columnspan=3)
    lbl2.grid(column=1, row=2, sticky="W")
    lbl3.grid(column=1, row=3, sticky="W")
    lbl4.grid(column=2, row=3, sticky="W")
    lbl5.grid(column=3, row=3, sticky="W")

    button.grid(column=2, row=2)

    window.mainloop()






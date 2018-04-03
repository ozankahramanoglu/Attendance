from gc import enable
from tkinter import *
from tkinter import filedialog
from tkinter import ttk

from pip._vendor.requests.packages.urllib3.packages import ssl_match_hostname
from xlrd import open_workbook, sheet


class fileReading:

    def __init__(self):
        print("init")

    sheet1 = sheet.Sheet

    def openingWorkbook(self):
        path = filedialog.askopenfilename(title="Select file")
        wb = open_workbook(path)
        first_sheet = wb.sheet_by_index(0)
        return first_sheet

    def getingSheet(self):
        global sheet1
        sheet1 = fileReading.openingWorkbook()



class data:

    def __init__(self):
        print("init")

    colnums = fileReading.sheet1.ncols
    rownums = fileReading.sheet1.nrows

    print(fileReading.sheet1.cell_value(rowx=0,colx=0))
    data = [[fileReading.sheet1.cell_value(r, c) for c in range(colnums)] for r in range(rownums)]
    del data[0]
    for i in range(rownums - 1):
        a = data[i][1].split(" ")
        j = ""
        for b in range(0, len(a) - 1):
            j = j + " " + a[b]

        h = a[-1] + ","

        data[i][1] = h + j
        print(data[i][1])

    listboxStudent = []

    for i in range(rownums - 1):
        listboxStudent.append(data[i][1] + ", " + str(data[i][0]))
        print(listboxStudent)



class GUI:

    def __init__(self):
        print("init")

    window = Tk()
    window.title("AttendanceKeeper V1.0")

    mainframe = ttk.Frame(window, padding="3 3 12 12")
    mainframe.grid(column=5, row=7, sticky=(N, W, E, S))

    lbl1 = Label(mainframe, text="Attendance Keeper V1.0", font=("Arial Bold", 25))
    lbl2 = Label(mainframe, text="Select students list Excel file", font=("Arial Bold", 8))
    lbl3 = Label(mainframe, text="Select students:", font=("Arial Bold", 8))
    lbl4 = Label(mainframe, text="Section:", font=("Arial Bold", 8))
    lbl5 = Label(mainframe, text="Attended Studets:", font=("Arial Bold", 8))

    lbl1.grid(column=2, row=1, sticky="N", columnspan=3)
    lbl2.grid(column=1, row=2, sticky="W", columnspan=2)
    lbl3.grid(column=1, row=3, sticky="W", columnspan=2)
    lbl4.grid(column=3, row=3, sticky="W")
    lbl5.grid(column=4, row=3, sticky="W", columnspan=2)

    studentList = ttk.Combobox(master=mainframe)
    studentList.grid(column=3, row=4)

    listbox1 = Listbox(master=mainframe)
    listbox1.grid(column=1, row=4, rowspan=3, columnspan=2)

    listbox2 = Listbox(master=mainframe)
    listbox2.grid(column=4, row=4, rowspan=3, columnspan=2)

    button1 = ttk.Button(master=mainframe, text="Import List", command=fileReading.getingSheet)
    button1.grid(column=3, row=2)

    button2 = ttk.Button(master=mainframe, text="Add =>", command=fileReading.getingSheet)
    button2.grid(column=3, row=5,)

    button3 = ttk.Button(master=mainframe, text="<= Remove", command=fileReading.getingSheet)
    button3.grid(column=3, row=6)

    lbl6 = Label(mainframe, text="Please select file type: ", font=("Arial Bold", 8))
    lbl6.grid(column=1, row=7, sticky="W")

    filetype = ttk.Combobox(master=mainframe)
    filetype.grid(column=2, row=7)

    lbl7 = Label(mainframe, text="Please Enter week: ", font=("Arial Bold", 8))
    lbl7.grid(column=3, row=7, sticky="W")

    entry = Entry(mainframe, width=30)
    entry.grid(column=4, row=7)

    button4 = ttk.Button(master=mainframe, text="Export as file", command=fileReading.getingSheet)
    button4.grid(column=5, row=7)

    window.mainloop()



from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from xlrd import open_workbook

class student:
    def __init__(self,ID,Name,Department,Section):
        studentName = Name
        studentId = ID
        studentDepartment = Department
        studentSection = Section


class studentList:

    def __init__(self,path):
        wb = open_workbook(path)
        first_sheet = wb.sheet_by_index(0)

        colnums = first_sheet.ncols
        rownums = first_sheet.nrows

        data = [[first_sheet.cell_value(r, c) for c in range(colnums)] for r in range(rownums)]

        del data[0]
        print(data)
        keys = []
        for l in range (len(data)):
            keys.append(data[l][3])
        set(keys)
        print(keys)
        studentDictionary = {"key":"value"}

        for i in data:
            for j in keys:
                print(j)
                if i[3] in keys:
                    print("ife girdi")
                    studentDictionary = [str(j)][len(studentDictionary[j])+1] : student(i[0],i[1],i[2],i[3])

        print(studentDictionary)



class GUI:

    path = ""

    def importExcel():
        global path
        path = filedialog.askopenfilename(title="Select file")

        listOfStudent = studentList(path)

    def __init__(self):
        window = Tk()
        window.title("AttendanceKeeper V1.0")

        mainframe = ttk.Frame(window, padding="3 3 12 12")
        mainframe.grid(column=5, row=7, sticky=(N, W, E, S))

        lbl1 = Label(mainframe, text="Attendance Keeper V1.0", font=("Arial Bold", 25))
        lbl2 = Label(mainframe, text="Select students list Excel file: ", font=("Arial Bold", 16))
        lbl3 = Label(mainframe, text="Select students:", font=("Arial Bold", 14))
        lbl4 = Label(mainframe, text="Section:", font=("Arial Bold", 14))
        lbl5 = Label(mainframe, text="Attended Students:", font=("Arial Bold", 14))
        lbl6 = Label(mainframe, text="Please select file type: ", font=("Arial Bold", 12))
        lbl7 = Label(mainframe, text="Please enter week: ", font=("Arial Bold", 12))

        lbl1.grid(column=2, row=1, sticky="N", columnspan=3)
        lbl2.grid(column=1, row=2, sticky="W", columnspan=2)
        lbl3.grid(column=1, row=3, sticky="W", columnspan=2)
        lbl4.grid(column=3, row=3, sticky="S")
        lbl5.grid(column=4, row=3, sticky="W", columnspan=2)
        lbl6.grid(column=1, row=7, sticky="W")
        lbl7.grid(column=3, row=7, sticky="E")

        studentList = ttk.Combobox(master=mainframe)
        studentList.grid(column=3, row=4, sticky="N")

        listbox1 = Listbox(master=mainframe)
        listbox1.grid(column=1, row=4, sticky=(S, E, W, N), rowspan=3, columnspan=2)

        listbox2 = Listbox(master=mainframe)
        listbox2.grid(column=4, row=4, sticky=(S, E, W, N), rowspan=3, columnspan=2)

        button1 = ttk.Button(master=mainframe, text="Import List", command=GUI.importExcel)
        button1.grid(column=3, row=2, sticky=(W, E))

        button2 = ttk.Button(master=mainframe, text="Add =>")
        button2.grid(column=3, row=4, sticky=(W, E, S))

        button3 = ttk.Button(master=mainframe, text="<= Remove")
        button3.grid(column=3, row=5, sticky=(W, E, N))

        button4 = ttk.Button(master=mainframe, text="Export as file")
        button4.grid(column=5, row=7)

        filetype = ttk.Combobox(master=mainframe)
        filetype.grid(column=2, row=7)

        entry = Entry(mainframe, width=30)
        entry.grid(column=4, row=7)

        window.mainloop()

if __name__ == "__main__":
    gui = GUI()

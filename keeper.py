from tkinter import *
from tkinter import filedialog
from tkinter import ttk
from xlrd import open_workbook
import xlwt

class student:

    def __init__(self, ID, Name, Department, Section):
        self.Name = Name
        self.ID = ID
        self.Department = Department
        self.Section = Section

    def __str__(self):
        return self.Name + "-" + self.ID + "-" + self.Department + "-" + self.Section

    def getName(self):
        return self.Name
    def getID(self):
        return self.ID
    def getDepartment(self):
        return self.Department
    def getSection(self):
        return self.Section



class studentList:

    def __init__(self, path):
        wb = open_workbook(path)
        first_sheet = wb.sheet_by_index(0)

        global colnums
        colnums = first_sheet.ncols
        global rownums
        rownums = first_sheet.nrows

        data = [[first_sheet.cell_value(r, c) for c in range(colnums)] for r in range(rownums)]

        del data[0]
        global studentArray
        studentArray = []
        keys = []

        for l in range(len(data)):
            keys.append(data[l][3])
        keys = sorted(set(keys))

        for i in range(len(keys)):
            studentArray.append([])
            for j in range(rownums - 1):
                if data[j][3] == keys[i]:
                    newStudent = student(ID=str(data[j][0]), Name=str(data[j][1]), Department=str(data[j][2]), Section=str(data[j][3]))
                    studentArray[i].append(newStudent)

    def getSectionList(self,section):
        selectedArray = []
        for k in range(len(studentArray[section])):
            text = (studentArray[section][k].Name).split(' ')
            tmptext = (text[-1] + ", ")
            text[-1] = ""
            for i in text:
                tmptext = tmptext + " " + i
            tmptext = tmptext + ", " + "{:.0f}".format(float(studentArray[section][k].ID))
            tmptext = tmptext + ";" + studentArray[section][k].Department
            selectedArray.append(tmptext)
            selectedArray = sorted(selectedArray)
        return selectedArray

    def getSection(self):
        sectionArray = []
        for k in range(len(studentArray)):
            for j in range(len(studentArray[k])):
                sectionArray.append(studentArray[k][j].Section)
        return sectionArray

    def getName(self):
        nameArray = []
        for k in range(len(studentArray)):
            for j in range(len(studentArray[k])):
                nameArray.append(studentArray[k][j].Name)
        return nameArray

    def getID(self):
        idArray = []
        for k in range(len(studentArray)):
            for j in range(len(studentArray[k])):
                idArray.append(studentArray[k][j].ID)
        return idArray

    def getDepartment(self):
        departmentArray = []
        for k in range(len(studentArray)):
            for j in range(len(studentArray[k])):
                departmentArray.append(studentArray[k][j].Department)
        return departmentArray


class GUI:

    comboValue = []
    stuList1 = []
    stuList2 = []

    def exportingFile(self):
        global tmp
        tmp = ""
        if filetype.current() == 0:
            name = entry.get() + ".txt"
            f = open(name, 'w+')
            for i in range(len(self.stuList2)):
                tmp = self.stuList2[i].split(";")
                tmp2 = tmp[1]
                tmp = tmp[0].split(", ")
                tmp.append(tmp2)
                f.write(tmp[2] + "\t" + tmp[1] + " " + tmp[0] + "\t" + tmp[3] + "\n")
            f.close()
        elif filetype.current() == 1:
            wb = xlwt.Workbook()
            ws = wb.add_sheet('Attendance')
            ws.write(0, 0, "ID")
            ws.write(0, 1, "NAME")
            ws.write(0, 2, "DEPARTMENT")
            for i in range(len(self.stuList2)):
                tmp = self.stuList2[i].split(";")
                tmp2 = tmp[1]
                tmp = tmp[0].split(", ")
                tmp.append(tmp2)
                ws.write(i+1, 0, tmp[2])
                ws.write(i+1, 1, tmp[1] + " " + tmp[0])
                ws.write(i+1, 2, tmp[3])
            wb.save(entry.get() + ".xls ")
        else:
            raise BaseException("File type is not supported")


    def updatelist3(self):
        for i in range(len(listbox2.curselection())):
            tmp = self.stuList1[listbox2.curselection()[i]].split(";")
            listbox1.insert(END, tmp[0])
        for i in range(len(listbox2.curselection())):
            tmp = listbox2.curselection()[0]
            listbox2.delete(tmp)

    def updatelist2(self):
        for i in range(len(listbox1.curselection())):
            tmp = self.stuList1[listbox1.curselection()[i]].split(";")
            listbox2.insert(END, tmp[0])
            self.stuList2.append(self.stuList1[listbox1.curselection()[i]])
        for i in range(len(listbox1.curselection())):
            tmp = listbox1.curselection()[0]
            listbox1.delete(tmp)


    def updatelist(self):
        self.stuList1 = azo.getSectionList(studentList1.current())
        listbox1.delete(0, END)
        for i in range(len(self.stuList1)):
            listbox1.insert(END, (self.stuList1[i]))


    def importExcel(self):
        path = filedialog.askopenfilename(title="Select file")
        global azo
        azo = studentList(path)
        self.comboValue = (sorted(set(list(azo.getSection()))))
        global studentList1,listbox1
        studentList1['values'] = self.comboValue
        studentList1.current(0)
        self.stuList1 = azo.getSectionList(studentList1.current())
        for i in range(len(self.stuList1)):
            tmp = self.stuList1[i].split(";")
            listbox1.insert(END, tmp[0])
        studentList1.grid(column=3, row=4, sticky="N")



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

        global studentList1, listbox1, listbox2
        studentList1 = ttk.Combobox(master=mainframe)
        studentList1.grid(column=3, row=4, sticky="N")
        studentList1.bind("<<ComboboxSelected>>", lambda e: self.updatelist())

        listbox1 = Listbox(master=mainframe,selectmode=MULTIPLE)
        #scrollbar1 = Scrollbar(listbox1)
        #scrollbar1.pack(side=RIGHT, fill=Y)
        listbox1.grid(column=1, row=4, sticky=(S, E, W, N), rowspan=3, columnspan=2)

        listbox2 = Listbox(master=mainframe,selectmode=MULTIPLE)
        #scrollbar2 = Scrollbar(listbox2)
        #scrollbar2.pack(side=RIGHT, fill=Y)
        listbox2.grid(column=4, row=4, sticky=(S, E, W, N), rowspan=3, columnspan=2)

        button1 = ttk.Button(master=mainframe, text="Import List", command=self.importExcel)
        button1.grid(column=3, row=2, sticky=(W, E))

        button2 = ttk.Button(master=mainframe, text="Add =>", command=self.updatelist2)
        button2.grid(column=3, row=4, sticky=(W, E, S))

        button3 = ttk.Button(master=mainframe, text="<= Remove", command=self.updatelist3)
        button3.grid(column=3, row=5, sticky=(W, E, N))

        button4 = ttk.Button(master=mainframe, text="Export as file", command=self.exportingFile)
        button4.grid(column=5, row=7)

        global filetype
        filetype = ttk.Combobox(master=mainframe)
        filetype['values'] = ("txt", "xls", "csv")
        filetype.current(0)
        filetype.grid(column=2, row=7)

        global entry
        entry = Entry(mainframe, width=30)
        entry.grid(column=4, row=7)

        window.mainloop()

if __name__ == "__main__":
    gui = GUI()


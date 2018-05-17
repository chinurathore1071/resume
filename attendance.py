from Tkinter import *
import xlrd,xlwt
from xlutils.copy import copy

"""
wb=xlwt.Workbook()
ws=wb.add_sheet("Studenst Name")
ws.write(0,0,'priyanka')
ws.write(1,0,'deepika')
ws.write(2,0,'sejal')

ws.write(0,1,10)
ws.write(1,1,11)
ws.write(2,1,12)

ws.write(0,2,'ece')
ws.write(1,2,'ece')
ws.write(2,2,'ece')



wb.save('newbook.xls')
"""

window=Tk()
window.title("attendance sheet")
window.geometry('400x400')
window.configure(background='orange')

Roll_number=Label(window,text="enter roll number")
Roll_number.pack(fill=X)

entry_roll = Entry(window)
entry_roll.pack(fill=X)

entry_name = Entry(window)
entry_name.pack(fill=X)


Value=Label(window,text="")

             
def clicked():
    roll = int ( entry_roll.get() ) - 1
    name = str(entry_name.get() )

    rb = xlrd.open_workbook("HP/students.xls")

    wb = copy(rb)

    sh = wb.get_sheet(0)

    sh.write(roll,0,name)

    wb.save("HP/students.xls")
    
    Value.configure(text="Value Edited")         
             
btn = Button(window, text="Press to  input in Excell",command=clicked)
btn.pack(fill=X)
window.mainloop()

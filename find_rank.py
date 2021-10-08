import tkinter
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox
import openpyxl
from openpyxl import load_workbook
import pyttsx3

window = Tk()
window.title('tm.py11')
window.geometry("900x600")

bangchucai = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']


# Them label
nhan = tkinter.Label(window, text = "TNNT21", fg = 'blue', font = ('Arial', 30))
nhan.grid(column = 0, row = 0)

sbd = tkinter.Label(window, text = "Số báo danh", fg = 'black', font = ('Arial', 10))
sbd.grid(column = 0, row = 1)

khoi = tkinter.Label(window, text = "Khối thi", fg = 'black', font = ('Arial', 10))
khoi.grid(column = 0, row = 2)

luuy = tkinter.Label(window, text = "Lưu ý: Kết quả dựa trên điểm thi TN (không tính điểm cộng) và chỉ mang tính tương đối.", fg = 'red', font = ('Arial, 10'))
luuy.grid(column = 7, row = 3)

# Them Textbox
txt = Entry(window, width = 20)
txt.grid(column = 3, row = 1)


combo = Combobox(window)
combo['value'] = ("A00", "A01", "B00", "C00", "D01", "D07")
combo.grid(column = 3, row = 2)


sbd_st = 45000000

def say_hello():
    robot_brain = "Hello, Welcom to Minh's tool"

    robot_mouth = pyttsx3.init()
    robot_mouth.say(robot_brain)
    robot_mouth.runAndWait()

def say_bye():
    robot_brain = "Good bye"

    robot_mouth = pyttsx3.init()
    robot_mouth.say(robot_brain)
    robot_mouth.runAndWait()

def get_value_excel(filename, cellname):
    wb = openpyxl.load_workbook(filename)
    Sheet1 = wb['Sheet1']
    wb.close()
    return Sheet1[cellname].value

def HandleButton1():
    tenkhoi = combo.get()
    if tenkhoi == 'A00':
        stt = 1
    elif tenkhoi == 'A01':
        stt = 2
    elif tenkhoi == 'B00':
        stt = 3
    elif tenkhoi == 'C00':
        stt = 4
    elif tenkhoi == 'D01':
        stt = 5
    else: stt = 6
    tmp = int(txt.get()) - sbd_st
    ten_o = bangchucai[stt] + str(tmp)
    rankk = get_value_excel('ttm.xlsx', ten_o)
    if txt.get() == "45004568":
        messagebox.showinfo("Error =)))", "Ai cho mà xem")
        messagebox.showinfo("Rank", "Rank của " + txt.get() + " ở khối " + combo.get() + " là: " + str(rankk))
        robot_mouth = pyttsx3.init()
        robot_mouth.say("hehe")
        robot_mouth.runAndWait()
    elif txt.get() == "45004569":
        messagebox.showinfo("Rank", "Tui là sanh diên rồi xem rank làm gì :)))")
    else: 
        messagebox.showinfo("Rank", "Rank của " + txt.get() + " ở khối " + combo.get() + " là: " + str(rankk))
    say_bye()
    return


say_hello()
but1 = tkinter.Button(window, text = "Submit", command = HandleButton1, fg = 'green')
but1.grid(column = 3, row = 3)

window.mainloop()

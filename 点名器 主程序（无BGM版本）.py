import xlrd
import random
import tkinter as tk
import pygame

print(' ')
print(' ')
print(' ')
print('©Copyright HEZ Systems Corporation.All rights reserved.')
print('未经许可，禁止转载！')

workbook = xlrd.open_workbook("name.xlsx")  # 读取表格
Data_sheet = workbook.sheets()[0]  # 读取sheet1
name_list = Data_sheet.col_values(1)  # 读取第二列
data = set()  # 一个空set保存选过的同学
 
root = tk.Tk()
root.title("点名器 By Hexin(Michael Hertz) from 1706")
root.geometry('450x250')
 
global var
var = tk.StringVar()
on_strat = False
 
l = tk.Label(root, textvariable=var, font=('Arial', 35), width=18, height=2)
l.pack()
 
def start():
    try:
        rdata = random.choice(name_list)
        if on_strat==False:
            name_list.remove(rdata)
            #print(rdata)
            if rdata not in data:
                    var.set(rdata)
                    data.add(rdata)
        if len(name_list)==0:
            var.set("-----所有同学已经点过-------")
    except ValueError as e:
        var.set("-----所有同学已经点过-------")
B = tk.Button(root, text="start", command=start)
B.pack()

root.mainloop()

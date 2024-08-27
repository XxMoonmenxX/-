import tkinter as tk
from tkinter import ttk
from openpyxl import load_workbook
import sqlite3 as sql
from datetime import date

con = sql.connect('test.txt')
cur = con.cursor()
soc = cur.fetchall()

win = tk.Tk()
win.geometry("200x200")
bg = tk.PhotoImage(file="logo.png")
img = ttk.Label(win, image=bg)
img.place(x=0, y=0)

canvas = tk.Canvas(win, borderwidth=0, background="#ffffff")
frame= tk.Frame(canvas, background="#ffffff")
vsb = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=vsb.set)
vsb.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True)
canvas.create_window((4, 4), window=frame, anchor="nw")

for i in range(20):
    lbl = tk.Label(frame, text=f"Label {i}")
    lbl.pack(side="top")

canvas.update_idletasks()
canvas.configure(scrollregion=canvas.bbox("all"))

win.mainloop()
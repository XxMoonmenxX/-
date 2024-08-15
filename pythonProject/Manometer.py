from tkinter import *
from tkinter import ttk
from openpyxl import load_workbook

from datetime import date

import sqlite3 as sql

con = sql.connect('test.txt')
cur = con.cursor()
soc = cur.fetchall()


def absolutepog():
    s = float(entry.get()) * float(entry2.get()) / 100
    label["text"] = s


def ne_znaiu():
    hui = str(float(entry3.get()) - float((entry4.get())))
    hui = str(round(float(hui), 5))
    label1["text"] = hui


def procent_govna():
    hui = str(float(entry3.get()) - float((entry4.get())))
    hui = str(round(float(hui), 5))
    loh = str(float(hui) / float(entry2.get()) * 100)
    loh = (round(float(loh), 2))
    label2["text"] = loh


root = Tk()
root.title("Рассчет погрешности")
root.geometry("768x1280")
root.resizable(width=False, height=False)
bg = PhotoImage(file="logo.png")
img = Label(root, image=bg)
img.place(x=0, y=0)
current_date = date.today()
"""
tab_control = ttk.Notebook(root)

vkladka1 = ttk.Frame(tab_control)
vkladka2 = ttk.Frame(tab_control)

tab_control.add(vkladka1, text='1')

bg = PhotoImage(file="logo.png")

img = Label(root, image=bg)
img.place(x=0, y=0)

tab_control.add(vkladka2, text='2')

bg = PhotoImage(file="logo.png")

img = Label(root, image=bg)
img.place(x=0, y=0)

lll = Label(vkladka1)
lll.grid(column=0, row=3)

lll1 = Label(vkladka2)
lll1.grid(column=1, row=4)

tab_control.pack(expand=1, fill='both')
"""
prikol1 = Menu(root)
root.config(menu=prikol1)

prikol2 = Menu(prikol1, tearoff=0)
prikol2.add_command(label='Закрой')

prikol1.add_cascade(label='Посмотри', menu=prikol2)

lbl = Label(root, text="Класс точности")
lbl.pack(anchor=NW, padx=6, pady=6)

entry = ttk.Entry()
entry.pack(anchor=NW, padx=6, pady=6)

lbl1 = Label(root, text="Диапазон")
lbl1.pack(anchor=NW, padx=6, pady=6)

entry2 = ttk.Entry()
entry2.pack(anchor=NW, padx=6, pady=6)

lbl11 = Label(root, text="СИ")
lbl11.pack(anchor=NW, padx=6, pady=6)

entry22 = ttk.Entry()
entry22.pack(anchor=NW, padx=6, pady=6)

btn = ttk.Button(text="Рассчет допустимой погрешности", command=absolutepog)
btn.pack(anchor=NW, padx=6, pady=6)

label = ttk.Label()
label.pack(anchor=NW, padx=6, pady=6)

lbl2 = Label(root, text="Оцифрованная точка")
lbl2.pack(anchor=NW, padx=6, pady=6)

entry3 = ttk.Entry()
entry3.pack(anchor=NW, padx=6, pady=6)

lbl3 = Label(root, text="Показания эталона")
lbl3.pack(anchor=NW, padx=6, pady=6)

entry4 = ttk.Entry()
entry4.pack(anchor=NW, padx=6, pady=6)

btn = ttk.Button(text="Рассчет разности шага", command=ne_znaiu)
btn.pack(anchor=NW, padx=6, pady=6)

label1 = ttk.Label()
label1.pack(anchor=NW, padx=6, pady=6)

btn = ttk.Button(text="Рассчет % погрешности шага", command=procent_govna)
btn.pack(anchor=NW, padx=6, pady=6)

label2 = ttk.Label()
label2.pack(anchor=NW, padx=6, pady=6)

lbl4 = Label(root, text="Название манометра")
lbl4.pack(anchor=NW, padx=6, pady=6)

entry5 = ttk.Entry()
entry5.pack(anchor=NW, padx=6, pady=6)

lbl5 = Label(root, text="Номер манометра")
lbl5.pack(anchor=NW, padx=6, pady=6)

entry6 = ttk.Entry()
entry6.pack(anchor=NW, padx=6, pady=6)




if  entry.get() < label2["text"]:
    k = 'Годен'
elif entry.get() > label2["text"]:
    k = 'Не годен'
def centrtxt():
    with con:
        print('Данные внесены ')
        # cur.execute("CREATE TABLE IF NOT EXISTS `test` (`name` STRING, `number` STRING, `kt` STRING, `diap` STRING, `k` STRING)")
        name = entry5.get()
        number = str(entry6.get())
        kt = str(entry.get())
        diap = str(entry2.get() + entry22.get())
        huii = str(float(entry3.get()) - float((entry4.get())))
        huii = str(round(float(huii), 5))
        lohh = str(float(huii) / float(entry2.get()) * 100)
        if entry.get() < lohh:
            k = 'Не годен'
        elif entry.get() > lohh:
            k = 'Годен'
        print(k)
        # cur.execute(f"INSERT INTO `test` VALUES ('{name}', '{number}', '{kt}', '{diap}', '{k}' )")
        xl = 'Журнал.xlsx'
        omg = load_workbook(xl)
        ogm = omg['Sheet1']
        ogm.append([current_date, name, number, kt, diap, k])
        omg.save(xl)
        omg.close()
        """rows = cur.fetchall()
         for row in rows:
         print(row)
         con.commit()
         cur.close()"""


centrtext = ttk.Button(text="Внести данные", command=centrtxt)
centrtext.pack(anchor=NW, padx=6, pady=6)

"""def centrtxt2():
cur.execute("SELECT * FROM `test`")
res = cur.fetchall()
for row in res:
print(row)"""


"""centrtext2 = ttk.Button(text="Вынести данные", command=centrtxt2)
centrtext2.pack(anchor=NW, padx=6, pady=6)"""

root.mainloop()

from tkinter import *
from tkinter import ttk
from openpyxl import load_workbook

from datetime import date

import sqlite3 as sql

con = sql.connect('test.txt')
cur = con.cursor()
soc = cur.fetchall()

#https://selectel.ru/blog/tutorials/tkinter-library-in-python/ полезная инструкция

def absolutepog():
    s = float(entry.get()) * float(entry2.get()) / 100
    label["text"] = s


def procent_govna():
    hui = str(float(entry3.get()) - float((entry4.get())))
    hui = str(round(float(hui), 5))
    hui = str(float(entry3.get()) - float((entry4.get())))
    hui = str(round(float(hui), 5))
    loh = str(float(hui) / float(entry2.get()) * 100)
    loh = (round(float(loh), 3))
    label2["text"] = loh

    huii = str(float(entryl3.get()) - float((entryl4.get())))
    huii = str(round(float(huii), 5))
    lohh = str(float(huii) / float(entry2.get()) * 100)
    lohh = (round(float(lohh), 3))
    labell2["text"] = lohh

    huiii = str(float(entryll3.get()) - float((entryll4.get())))
    huiii = str(round(float(huiii), 5))
    lohhh = str(float(huiii) / float(entry2.get()) * 100)
    lohhh = (round(float(lohhh), 3))
    labelll2["text"] = lohhh

    huiiii = str(float(entryllll3.get()) - float((entrylll4.get())))
    huiiii = str(round(float(huiiii), 5))
    lohhhh = str(float(huiiii) / float(entry2.get()) * 100)
    lohhhh = (round(float(lohhhh), 3))
    labellll2["text"] = lohhhh

    huiiiii = str(float(entrylllll3.get()) - float((entryllll4.get())))
    huiiiii = str(round(float(huiiiii), 5))
    lohhhhh = str(float(huiiiii) / float(entry2.get()) * 100)
    lohhhhh = (round(float(lohhhhh), 3))
    labelllll2["text"] = lohhhhh

    huiiiiii = str(float(entryllllll3.get()) - float((entrylllll4.get())))
    huiiiiii = str(round(float(huiiiiii), 5))
    lohhhhhh = str(float(huiiiiii) / float(entry2.get()) * 100)
    lohhhhhh = (round(float(lohhhhhh), 3))
    labellllll2["text"] = lohhhhhh


root = Tk()
root.title("Рассчет погрешности")
root.geometry("768x1280")
#root.resizable(width=False, height=False)
bg = PhotoImage(file="logo.png")
img = Label(root, image=bg)
img.place(x=0, y=0)
current_date = date.today()

tab_control = ttk.Notebook(root)

vkladka1 = ttk.Frame(tab_control)
vkladka2 = ttk.Frame(tab_control)

vkladka1.pack(fill=BOTH, expand=True)
vkladka2.pack(fill=BOTH, expand=True)

tab_control.pack(expand=0, fill='both')



tab_control.add(vkladka1, text='Данные манометра')



tab_control.add(vkladka2, text='Рассчет погрешности')


lbl = Label(root, text="Класс точности")
lbl.pack(anchor=NW)

entry = ttk.Entry()
entry.pack(anchor=NW)

lbl1 = Label(root, text="Диапазон")
lbl1.pack(anchor=NW)

entry2 = ttk.Entry()
entry2.pack(anchor=NW)

lbl11 = Label(root, text="СИ")
lbl11.pack(anchor=NW)

entry22 = ttk.Entry()
entry22.pack(anchor=NW)

btn = ttk.Button(text="Рассчет допустимой погрешности", command=absolutepog)
btn.pack(anchor=NW)

label = ttk.Label()
label.pack(anchor=NW)

lbl2 = Label(root, text="Оцифрованная точка")
lbl2.pack(anchor=NW)

entry3 = ttk.Entry()
entry3.pack(anchor=NW)

lbl3 = Label(root, text="Показания эталона")
lbl3.pack(anchor=NW)

entry4 = ttk.Entry()
entry4.pack(anchor=NW)

lbll2 = Label(root, text="Оцифрованная точка 2")
lbll2.pack(anchor=NW)

entryl3 = ttk.Entry()
entryl3.pack(anchor=NW)

lbll3 = Label(root, text="Показания эталона 2")
lbll3.pack(anchor=NW)

entryl4 = ttk.Entry()
entryl4.pack(anchor=NW)

lblll2 = Label(root, text="Оцифрованная точка 3")
lblll2.pack(anchor=NW)

entryll3 = ttk.Entry()
entryll3.pack(anchor=NW)

lblll3 = Label(root, text="Показания эталона 3")
lblll3.pack(anchor=NW)

entryll4 = ttk.Entry()
entryll4.pack(anchor=NW)

lbllll2 = Label(root, text="Оцифрованная точка 4")
lbllll2.pack(anchor=NW)

entryllll3 = ttk.Entry()
entryllll3.pack(anchor=NW)

lbllll3 = Label(root, text="Показания эталона 4")
lbllll3.pack(anchor=NW)

entrylll4 = ttk.Entry()
entrylll4.pack(anchor=NW)

lblllll2 = Label(root, text="Оцифрованная точка 5")
lblllll2.pack(anchor=NW)

entrylllll3 = ttk.Entry()
entrylllll3.pack(anchor=NW)

lblllll3 = Label(root, text="Показания эталона 5")
lblllll3.pack(anchor=NW)

entryllll4 = ttk.Entry()
entryllll4.pack(anchor=NW)

lbllllll2 = Label(root, text="Оцифрованная точка 6")
lbllllll2.pack(anchor=NW)

entryllllll3 = ttk.Entry()
entryllllll3.pack(anchor=NW)

lbllllll3 = Label(root, text="Показания эталона 6")
lbllllll3.pack(anchor=NW)

entrylllll4 = ttk.Entry()
entrylllll4.pack(anchor=NW)

#btn = ttk.Button(text="Рассчет разности шага", command=ne_znaiu)
#btn.pack(anchor=W, padx=6, pady=6)

btn = ttk.Button(text="Рассчет % погрешности шага", command=procent_govna)
btn.pack(anchor=W)

label2 = ttk.Label()
label2.pack(anchor=W)
labell2 = ttk.Label()
labell2.pack(anchor=NW)
labelll2 = ttk.Label()
labelll2.pack(anchor=NW)
labellll2 = ttk.Label()
labellll2.pack(anchor=NW)
labelllll2 = ttk.Label()
labelllll2.pack(anchor=NW)
labellllll2 = ttk.Label()
labellllll2.pack(anchor=NW)

lbl4 = Label(root, text="Название манометра")
lbl4.pack(anchor=NW)

entry5 = ttk.Entry()
entry5.pack(anchor=NW)

lbl5 = Label(root, text="Номер манометра")
lbl5.pack(anchor=NW)

entry6 = ttk.Entry()
entry6.pack(anchor=NW)

if entry.get() < label2["text"] or entry.get() < labell2["text"] or entry.get() < labelll2["text"] or entry.get() < \
        labellll2["text"] or entry.get() < labelllll2["text"] or entry.get() < labellllll2["text"]:
    k = 'Годен'
elif entry.get() > label2["text"] or entry.get() > labell2["text"] or entry.get() > labelll2["text"] or entry.get() > \
        labellll2["text"] or entry.get() > labelllll2["text"] or entry.get() < labellllll2["text"]:
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
centrtext.pack(anchor=NW)

root.mainloop()

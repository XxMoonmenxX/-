import tkinter as tk
from tkinter import ttk
from openpyxl import load_workbook
import sqlite3 as sql
from datetime import date

con = sql.connect('test.txt')
cur = con.cursor()
soc = cur.fetchall()

win = tk.Tk()
win.geometry("480x840")
bg = tk.PhotoImage(file="logo.png")
img = ttk.Label(win, image=bg)
img.place(x=0, y=0)




def oncontextaction(event):
    name_of_x_y = nb.identify(event.x, event.y)
    if name_of_x_y:
        x = event.x
        if 10 <= x < 177:
            index = 0
            print(f'ПКМ:  {nb.tab(index)["text"]}; index = {index}')  #
        if 177 <= x < 342:
            index = 1
            print(f'ПКМ:  {nb.tab(index)["text"]}; index = {index}')
        if 342 <= x < 508:
            index = 2
            print(f'ПКМ:  {nb.tab(index)["text"]}; index = {index}')

def absolutepog():
    s = float(entry.get()) * float(entry2.get()) / 100


def procent_govna():
    hui = str(float(entry3.get()) - float((entry4.get())))
    hui = str(round(float(hui), 5))
    hui = str(float(entry3.get()) - float((entry4.get())))
    hui = str(round(float(hui), 5))
    loh = str(float(hui) / float(entry2.get()) * 100)
    loh = (round(float(loh), 3))
    label["text"] = loh

    huii = str(float(entryl3.get()) - float((entryl4.get())))
    huii = str(round(float(huii), 5))
    lohh = str(float(huii) / float(entry2.get()) * 100)
    lohh = (round(float(lohh), 3))
    label2["text"] = lohh

    huiii = str(float(entryll3.get()) - float((entryll4.get())))
    huiii = str(round(float(huiii), 5))
    lohhh = str(float(huiii) / float(entry2.get()) * 100)
    lohhh = (round(float(lohhh), 3))
    label3["text"] = lohhh

    huiiii = str(float(entryllll3.get()) - float((entrylll4.get())))
    huiiii = str(round(float(huiiii), 5))
    lohhhh = str(float(huiiii) / float(entry2.get()) * 100)
    lohhhh = (round(float(lohhhh), 3))
    label4["text"] = lohhhh

    huiiiii = str(float(entrylllll3.get()) - float((entryllll4.get())))
    huiiiii = str(round(float(huiiiii), 5))
    lohhhhh = str(float(huiiiii) / float(entry2.get()) * 100)
    lohhhhh = (round(float(lohhhhh), 3))
    label5["text"] = lohhhhh

    huiiiiii = str(float(entryllllll3.get()) - float((entrylllll4.get())))
    huiiiiii = str(round(float(huiiiiii), 5))
    lohhhhhh = str(float(huiiiiii) / float(entry2.get()) * 100)
    lohhhhhh = (round(float(lohhhhhh), 3))
    label6["text"] = lohhhhhh


color = '#21252b'
win.configure(background=color)
current_date = date.today()
"""
sky_color = "sky blue"
gold_color = "gold"
color_tab = "#ccdee0" 

# style
style = ttk.Style()
style.theme_create("beautiful", parent="alt", settings={
    "TNotebook": {
        "configure": {"tabmargins": [10, 10, 20, 10], "background": sky_color}},
    "TNotebook.Tab": {
        "configure": {"padding": [30, 15],
                      "background": sky_color,
                      "font": ('consolas italic', 14),

                      "width": 10,

                      "borderwidth": [3]},

        "map": {"background": [("selected", gold_color), ('!active', sky_color), ('active', color_tab)],
                "expand": [("selected", [1, 1, 1, 0])]}}})
style.theme_use("beautiful")
style.layout("Tab",
             [('Notebook.tab', {'sticky': 'nswe', 'children':
                 [('Notebook.padding', {'side': 'top', 'sticky': 'nswe', 'children':
                 # [('Notebook.focus', {'side': 'top', 'sticky': 'nswe', 'children':
                     [('Notebook.label', {'side': 'top', 'sticky': ''})],
                                        # })],
                                        })],
                                })]
             )
style.configure('TLabel', background=color, foreground='white')
style.configure('TFrame', background=color)"""

nb = ttk.Notebook(win, width=300, height=300)



fr1 = ttk.Frame(nb)
fr2 = ttk.Frame(nb)
fr3 = ttk.Frame(nb)

bg = tk.PhotoImage(file="logo.png")
img = ttk.Label(fr1, image=bg)
img.place(x=0, y=0)

bg1 = tk.PhotoImage(file="logo.png")
img1 = ttk.Label(fr2, image=bg1)
img1.place(x=0, y=0)

bg2 = tk.PhotoImage(file="logo.png")
img2 = ttk.Label(fr3, image=bg2)
img2.place(x=0, y=0)


lb1 = ttk.Label(fr1, text="Название манометра")
lb1.grid(column=0, row=0)

entry5 = ttk.Entry(fr1)
entry5.grid(column=0, row=1)

lb1 = ttk.Label(fr1, text="Номер манометра")
lb1.grid(column=0, row=2)

entry6 = ttk.Entry(fr1)
entry6.grid(column=0, row=3)

lb1 = ttk.Label(fr1, text="Класс точности")
lb1.grid(column=0, row=4)

entry = ttk.Entry(fr1)
entry.grid(column=0, row=5)

lb1 = ttk.Label(fr1, text="Диапазон")
lb1.grid(column=0, row=6)

entry2 = ttk.Entry(fr1)
entry2.grid(column=0, row=7)

lb1 = ttk.Label(fr1, text="Система измерения")
lb1.grid(column=0, row=8)

lb1 = ttk.Label(fr1)
lb1.grid(column=0, row=9)

entry22 = ttk.Entry(fr1)
entry22.grid(column=0, row=10)




canvas = tk.Canvas(fr2, borderwidth=0, background="#ffffff")
vsb = ttk.Scrollbar(fr2, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=vsb.set)
vsb.pack(side="right", fill="y")
canvas.create_window((1, 12), window=fr2, anchor="nw")
canvas.update_idletasks()
canvas.configure(scrollregion=canvas.bbox("all"))

lb2 = ttk.Label(fr2, text="Оцифрованная точка")
lb2.pack(padx=5, pady=3)

entry3 = ttk.Entry(fr2)
entry3.pack(padx=5, pady=3)

lb2 = ttk.Label(fr2, text="Показания эталона")
lb2.pack(padx=5, pady=3)

entry4 = ttk.Entry(fr2)
entry4.pack(padx=5, pady=3)

lb2 = ttk.Label(fr2, text="Оцифрованная точка 2")
lb2.pack(padx=5, pady=3)

entryl3 = ttk.Entry(fr2)
entryl3.pack(padx=5, pady=3)

lb2 = ttk.Label(fr2, text="Показания эталона 2")
lb2.pack(padx=5, pady=3)

entryl4 = ttk.Entry(fr2)
entryl4.pack(padx=5, pady=3)

lb2 = ttk.Label(fr2, text="Оцифрованная точка 3")
lb2.pack(padx=5, pady=3)

entryll3 = ttk.Entry(fr2)
entryll3.pack(padx=5, pady=3)

lb2 = ttk.Label(fr2, text="Показания эталона 3")
lb2.pack(padx=5, pady=3)

entryll4 = ttk.Entry(fr2)
entryll4.pack(padx=5, pady=3)

lb2 = ttk.Label(fr2, text="Оцифрованная точка 4")
lb2.pack(padx=5, pady=3)

entryllll3 = ttk.Entry(fr2)
entryllll3.pack(padx=5, pady=3)

lb2 = ttk.Label(fr2, text="Показания эталона 4")
lb2.pack(padx=5, pady=3)

entrylll4 = ttk.Entry(fr2)
entrylll4.pack(padx=5, pady=3)

lb2 = ttk.Label(fr2, text="Оцифрованная точка 5")
lb2.pack(padx=5, pady=3)

entrylllll3 = ttk.Entry(fr2)
entrylllll3.pack(padx=5, pady=3)

lb2 = ttk.Label(fr2, text="Показания эталона 5")
lb2.pack(padx=5, pady=3)

entryllll4 = ttk.Entry(fr2)
entryllll4.pack(padx=5, pady=3)

lb2 = ttk.Label(fr2, text="Оцифрованная точка 6")
lb2.pack(padx=5, pady=3)

entryllllll3 = ttk.Entry(fr2)
entryllllll3.pack(padx=5, pady=3)

lb2 = ttk.Label(fr2, text="Показания эталона 6")
lb2.pack(padx=5, pady=3)

entrylllll4 = ttk.Entry(fr2)
entrylllll4.pack(padx=5, pady=3)


#lb3 = ttk.Label(fr3, text="Tab3")

label = ttk.Label(fr3, text=' 1 точка')
label.pack(padx=5, pady=5)
label2 = ttk.Label(fr3, text=' 2 точка')
label2.pack(padx=5, pady=5)
label3 = ttk.Label(fr3,text=' 3 точка')
label3.pack(padx=5, pady=5)
label4 = ttk.Label(fr3,text=' 4 точка')
label4.pack(padx=5, pady=5)
label5 = ttk.Label(fr3,text=' 5 точка')
label5.pack(padx=5, pady=5)
label6 = ttk.Label(fr3,text=' 6 точка')
label6.pack(padx=5, pady=5)

lb1.grid(column=0, row=0)
lb2.pack(padx=5, pady=5)
#lb3.pack(padx=5, pady=5)
fr1.pack(padx=5, pady=5)
fr2.pack(padx=5, pady=5)
fr3.pack(padx=5, pady=5)

nb.add(fr1, text="Данные манометра")
nb.add(fr2, text="Рассчет погрешности манометра")
nb.add(fr3, text="Получение результата")

nb.pack(fill="both", expand=1, padx=0, pady=0)
nb.enable_traversal()

nb.bind("<Button-3>", oncontextaction)

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

        hui = str(float(entry3.get()) - float((entry4.get())))
        hui = str(round(float(hui), 5))
        hui = str(float(entry3.get()) - float((entry4.get())))
        hui = str(round(float(hui), 5))
        loh = str(float(hui) / float(entry2.get()) * 100)
        loh = (round(float(loh), 3))

        huii = str(float(entryl3.get()) - float((entryl4.get())))
        huii = str(round(float(huii), 5))
        lohh = str(float(huii) / float(entry2.get()) * 100)
        lohh = (round(float(lohh), 3))

        huiii = str(float(entryll3.get()) - float((entryll4.get())))
        huiii = str(round(float(huiii), 5))
        lohhh = str(float(huiii) / float(entry2.get()) * 100)
        lohhh = (round(float(lohhh), 3))

        huiiii = str(float(entryllll3.get()) - float((entrylll4.get())))
        huiiii = str(round(float(huiiii), 5))
        lohhhh = str(float(huiiii) / float(entry2.get()) * 100)
        lohhhh = (round(float(lohhhh), 3))

        huiiiii = str(float(entrylllll3.get()) - float((entryllll4.get())))
        huiiiii = str(round(float(huiiiii), 5))
        lohhhhh = str(float(huiiiii) / float(entry2.get()) * 100)
        lohhhhh = (round(float(lohhhhh), 3))

        huiiiiii = str(float(entryllllll3.get()) - float((entrylllll4.get())))
        huiiiiii = str(round(float(huiiiiii), 5))
        lohhhhhh = str(float(huiiiiii) / float(entry2.get()) * 100)
        lohhhhhh = (round(float(lohhhhhh), 3))

        if entry.get() < str(float(loh)) or entry.get() < str(float(lohh)) or entry.get() < str(float(lohhh)) or entry.get() < str(float(lohhhh)) or entry.get() < str(float(lohhhhh)) or entry.get() < str(float(lohhhhhh)):
            k = 'Не годен'
        elif entry.get() > str(float(loh)) or entry.get() > str(float(lohh)) or entry.get() > str(float(lohhh)) or entry.get() > str(float(lohhhh)) or entry.get() > str(float(lohhhhh)) or entry.get() > str(float(lohhhhhh)):
            k = 'Годен'
        print(k)
        # cur.execute(f"INSERT INTO `test` VALUES ('{name}', '{number}', '{kt}', '{diap}', '{k}' )")
        xl = 'Журнал.xlsx'
        omg = load_workbook(xl)
        ogm = omg['Лист1']
        ogm.append([current_date, name, number, kt, diap, k])
        omg.save(xl)
        omg.close()

btn = ttk.Button(fr3,text="Рассчет % погрешности шага", command=procent_govna)
btn.pack(padx=5, pady=5)
centrtext = ttk.Button(fr3,text="Внести данные", command=centrtxt)
centrtext.pack()

win.mainloop()
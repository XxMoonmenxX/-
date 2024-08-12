from tkinter import *
from tkinter import ttk



def absolutepog():
    s = float(entry.get()) * float(entry2.get()) / 100
    label["text"] = s
def ne_znaiu():
    hui = str(float(entry3.get()) - float((entry4.get() )))
    hui = str(round(float(hui), 5))
    label1 ["text"] = hui

def procent_govna():
    hui = str(float(entry3.get()) - float((entry4.get())))
    hui = str(round(float(hui), 5))
    loh = str(float(hui) / float(entry2.get()) * 100)
    loh = (round(float(loh), 2))
    label2 ["text"] = loh

root = Tk()
root.title("Рассчет погрешности")
root.geometry("240x360")

entry = ttk.Entry()
entry.pack(anchor=NW, padx=6, pady=6)

entry2 = ttk.Entry()
entry2.pack(anchor=NW, padx=6, pady=12)

btn = ttk.Button(text="Рассчет допустимой погрешности", command=absolutepog)
btn.pack(anchor=NW, padx=6, pady=6)

label = ttk.Label()
label.pack(anchor=NW, padx=6, pady=6)

entry3 = ttk.Entry()
entry3.pack(anchor=NW, padx=6, pady=6)

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

root.mainloop()
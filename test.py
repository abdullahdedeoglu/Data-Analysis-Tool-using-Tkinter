import tkinter as tk
from tkinter import *
from tkinter import ttk, filedialog
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import math

win = Tk()

win.geometry("700x700")
#win.attributes('-fullscreen', True)
# win.overrideredirect(True)
# win.geometry("{0}x{1}+0+0".format(win.winfo_screenwidth(),
#              win.winfo_screenheight()))

style = ttk.Style()
style.theme_use('clam')

frame = Frame(win)
frame.pack(pady=20)


dataset = pd.read_excel(
    r'C:\Users\abdul\Desktop\Matlab Calismacalar\20010011046\arrhythmia.xlsx')


def missvalue(tree):

    dfColVal = []
    for i in dataset.columns.values:
        dfColVal = dataset[i].tolist()
        sum = 0
        for j in dfColVal:
            if (j != "?"):
                sum += j
        colavg = int(sum/len(dfColVal))
        dataset[i] = dataset[i].replace(['?'], colavg)
        dfColVal.clear()

    clear_treeview(tree)
    open_file(tree)


def ortalama(dataset, x):
    dfColVal = []
    dfColVal = dataset[dataset.columns[int(x)-1]].tolist()
    sum = 0
    for i in dfColVal:
        sum += i
    colavg = sum/len(dfColVal)
    return colavg


def medyan(dataset, x):
    dfColVal = []
    dfColVal = dataset[dataset.columns[int(x)-1]].tolist()
    dfColVal.sort()
    if len(dfColVal) % 2 == 0:
        medyan1 = dfColVal[len(dfColVal)//2]
        medyan2 = dfColVal[len(dfColVal)//2-1]
        meydan = (medyan1 + medyan2)/2
    else:
        meydan = dfColVal[len(dfColVal)//2]
    return meydan


def modbul(dataset, x):
    dfColVal = []
    dfColVal = dataset[dataset.columns[int(x)-1]].tolist()
    frekans = {}
    for i in dfColVal:
        frekans.setdefault(i, 0)
        frekans[i] += 1
    maks = max(frekans.values())
    modx = []
    for i, j in frekans.items():
        if j == maks:
            modx.append(i)
    return modx


def frrekans(dataset, x):
    dfColVal = []
    dfColVal = dataset[dataset.columns[int(x)-1]].tolist()
    frekans = {}
    for i in dfColVal:
        frekans.setdefault(i, 0)
        frekans[i] += 1
    return frekans


def frekanstablo(frekans):
    plt.bar(list(frekans.keys()), frekans.values(), color='g')
    plt.show()


def qbir(dataset, x):
    dfColVal = []
    dfColVal = dataset[dataset.columns[int(x)-1]].tolist()
    dfColVal.sort()
    n = (len(dfColVal)+1)/4
    int_n = int(n)
    if (n % 1 != 0):
        lowervalue = dfColVal[int_n-1]
        uppervalue = dfColVal[int_n]
        return (lowervalue+uppervalue)/2
    return dfColVal[int_n-1]


def quc(dataset, x):
    dfColVal = []
    dfColVal = dataset[dataset.columns[int(x)-1]].tolist()
    dfColVal.sort()
    n = 3*(len(dfColVal)+1)/4
    int_n = int(n)
    if (n % 1 != 0):
        lowervalue = dfColVal[int_n-1]
        uppervalue = dfColVal[int_n]
        return (lowervalue+uppervalue)/2
    return dfColVal[int_n-1]


def aykiri(iqr, q1, q3, dataset, x):
    dfColVal = []
    dfColVal = dataset[dataset.columns[int(x)-1]].tolist()
    aykiri_degerler = []
    alt_sinir = q1-(1.5*iqr)
    ust_sinir = q3+(1.5*iqr)
    print(alt_sinir, ust_sinir)
    for i in dfColVal:
        if (i < alt_sinir):
            aykiri_degerler.append(i)
        elif (i > ust_sinir):
            aykiri_degerler.append(i)
    return aykiri_degerler


def bes_sayi(dataset, x, q1, q3, meydan):
    dfColVal = []
    dfColVal = dataset[dataset.columns[int(x)-1]].tolist()
    minsayi = min(dfColVal)
    maxsayi = max(dfColVal)
    Label(win, text="Min:"+str(minsayi)+"\t"+"Max:"+str(maxsayi)+"\t"+"Q1:" +
          str(q1)+"\t"+"Q3:"+str(q3)+"\t"+"Medyan:"+str(meydan),
          font=('Century 8 bold')).pack(pady=5)


def var_ve_ss(dataset, x):
    dfColVal = []
    dfColVal = dataset[dataset.columns[int(x)-1]].tolist()
    ort = ortalama(dataset, x)
    varyans = sum(pow(i-ort, 2) for i in dfColVal) / len(dfColVal)
    ss = math.sqrt(varyans)
    Label(win, text="Varyans:"+str(varyans)+"\t"+"Standart Sapma:"+str(ss)+"\t",
          font=('Century 8 bold')).pack(pady=5)


def open_file(tree):
    df = dataset

    tree["column"] = list(df.columns)
    tree["show"] = "headings"

    for col in list(df.columns):
        tree.column("#" + str(df.columns.get_loc(col)),
                    anchor=CENTER, stretch=NO, width=50)
        tree.heading("#" + str(df.columns.get_loc(col)), text=col)

    df_rows = df.to_numpy().tolist()
    for row in df_rows:
        tree.insert("", "end", values=row)

    tree.pack()


def clear_treeview(tree):
    tree.delete(*tree.get_children())


def getSummary(entry1):
    x = int(entry1.get())
    colavg = ortalama(dataset, x)
    Label(win, text="Sütun Ortalaması: " + str(colavg),
          font=('Century 8 bold')).pack(pady=5)
    meydan = medyan(dataset, x)
    Label(win, text="Medyan: " + str(meydan),
          font=('Century 8 bold')).pack(pady=5)
    modx = modbul(dataset, x)
    Label(win, text="Mod: " + str(modx),
          font=('Century 8 bold')).pack(pady=5)
    frekans = frrekans(dataset, x)
    Label(win, text="Frekans: " + str(frekans),
          font=('Century 6 bold')).pack(pady=5)
    dfColVal = []
    dfColVal = dataset[dataset.columns[int(x)-1]].tolist()
    dfColVal.sort()
    q1 = qbir(dataset, x)
    q3 = quc(dataset, x)
    iqr = q3-q1
    Label(win, text="IQR: " + str(iqr),
          font=('Century 8 bold')).pack(pady=5)
    aykiri_degerler = aykiri(iqr, q1, q3, dataset, x)
    Label(win, text="Aykırı degerler: " + str(aykiri_degerler),
          font=('Century 8 bold')).pack(pady=5)
    bes_sayi(dataset, x, q1, q3, meydan)
    var_ve_ss(dataset, x)
    plt.boxplot(dfColVal)
    plt.show()
    frekanstablo(frekans)


def main():
    tree = ttk.Treeview(frame)
    win.resizable(width=0, height=0)

    horScrllBar = ttk.Scrollbar(win, orient="horizontal", command=tree.xview)

    horScrllBar.pack(side='right', fill='x')
    horScrllBar.place(x=0, y=250, width=700)

    tree.configure(xscrollcommand=horScrllBar.set)

    entry1 = Entry(win, width=40)
    entry1.focus_set()
    entry1.pack()

    btnSummary = Button(win, text='Veri Özeti', bd='5',
                        command=lambda: getSummary(entry1))
    btnSummary.place(x=500, y=275)

    btn = Button(win, text='Eksik Veri Temizle', bd='5',
                 command=lambda: missvalue(tree))
    btn.place(x=575, y=275)
    m = Menu(win)
    win.config(menu=m)

    file_menu = Menu(m, tearoff=False)
    m.add_cascade(label="Menu", menu=file_menu)
    file_menu.add_command(label="Open Spreadsheet",
                          command=lambda: open_file(tree))
    label = Label(win, text='')
    label.pack(pady=20)

    win.mainloop()


main()

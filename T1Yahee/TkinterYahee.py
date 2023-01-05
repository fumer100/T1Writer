import tkinter as tk
import yaheeFunctions as m
import pandas as pd



hauptFenster = tk.Tk()
hauptFenster.winfo_toplevel().title('YAHEE T1')
instructionLabel = tk.Label(hauptFenster, text="Geben sie die Daten ein!")

def deleteInput():
    Sdus.delete(0,'end')
    Schiff.delete(0, 'end')
    BL.delete(0, 'end')
    BLDATUM.delete(0, 'end')
    rechnungsnr.delete(0, 'end')
    rechnungsdatum.delete(0, 'end')
    Containernr.delete(0, 'end')
    Incoterm.delete(0, 'end')
    Transportpreis.delete(0, 'end')

def createFile(Sdus,Schiff,BL,BLDATUM,rechnungsnr,rechnungsdatum,Containernr,Incoterm,Transportpreis,Inlandspreis):
    Sendungsnr = Sdus.get()
    Schiff1 = Schiff.get()
    BL1 = BL.get()
    BLDatum = float(BLDATUM.get())
    Rechnungsnr = rechnungsnr.get()
    Rechnungsdatum = float(rechnungsdatum.get())
    Containernr1 = Containernr.get()
    Incoterm1 = "FOB "+Incoterm.get()
    Transportpreis1 = Transportpreis.get()+" EUR"
    Inlandpreis1 =str(Inlandspreis)+" EUR"
    writer = pd.ExcelWriter(Sendungsnr + ".xlsx", engine='xlsxwriter')
    T1, verzollung = m.createT1('test.xlsx')
    m.createWorkbook(T1, verzollung, Sendungsnr, Schiff1, BL1, BLDatum ,Rechnungsnr,Rechnungsdatum,Containernr1, Incoterm1, Transportpreis1, Inlandpreis1,writer)
    print('done')
    deleteInput()
    return




Sdus = tk.Entry(hauptFenster)
Sdus.grid(row=0,column=1)
labelSdus=tk.Label(hauptFenster,text='SDUS').grid(row=0)



Schiff = tk.Entry(hauptFenster,)
Schiff.grid(row=1,column=1)
labelSchiff=tk.Label(hauptFenster,text='Schiff').grid(row=1)


BL = tk.Entry(hauptFenster)
BL.grid(row=2,column=1)
labelBL=tk.Label(hauptFenster,text='BL').grid(row=2)


BLDATUM = tk.Entry(hauptFenster)
BLDATUM.grid(row=3,column=1)
labelBLDATUM=tk.Label(hauptFenster,text='BLDatum').grid(row=3)


rechnungsnr = tk.Entry(hauptFenster)
rechnungsnr.grid(row=4,column=1)
labelRechnungsnr=tk.Label(hauptFenster,text='Rechnungsnr.').grid(row=4)


rechnungsdatum = tk.Entry(hauptFenster)
rechnungsdatum.grid(row=5,column=1)
labelrechnungsdatum=tk.Label(hauptFenster,text='Rechnungsdatum').grid(row=5)


Containernr = tk.Entry(hauptFenster)
Containernr.grid(row=6,column=1)
labelContainernr=tk.Label(hauptFenster,text='Containernr.').grid(row=6)


Incoterm = tk.Entry(hauptFenster)
Incoterm.grid(row=7,column=1)
labelIncoterm=tk.Label(hauptFenster,text='Incoterm').grid(row=7)


Transportpreis = tk.Entry(hauptFenster)
Transportpreis.grid(row=8,column=1)
labelTransportpreis=tk.Label(hauptFenster,text='Transportpreis').grid(row=8)


Inlandspreis=795

SdusButton = tk.Button(text='Create', command=lambda: createFile(Sdus,Schiff,BL,BLDATUM,rechnungsnr,rechnungsdatum,Containernr,Incoterm,Transportpreis,Inlandspreis))
SdusButton.grid(row=9,column=1)



hauptFenster.mainloop()


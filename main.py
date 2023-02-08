import pandas as pd
import numpy as np
from xlsxwriter.utility import xl_rowcol_to_cell
import tkinter as tki
from tkinter import *
from tkinter import filedialog
from pandastable import Table, TableModel


def sortfile(file):
    df = pd.read_excel(file)
    df = df.sort_values('ID')
    writer = pd.ExcelWriter(file)
    df.to_excel(writer, index=False)
    writer.close()

def comparefile():
    file1 = textFile1.get()
    file2 = textFile2.get()
    template = pd.read_excel(file1, na_values="np.nan", header=None)
    testSheet = pd.read_excel(file2, na_values=np.nan, header=None)

    rt, ct = template.shape
    rtest, ctest = testSheet.shape

    df = pd.DataFrame(columns=['Cell_Location', 'file1', 'file2'])

    for rowNo in range(max(rt, rtest)):
        for colNo in range(max(ct, ctest)):
            # Fetching the template value at a cell
            try:
                template_val = template.iloc[rowNo, colNo]
            except:
                template_val = np.nan

            # Fetching the testsheet value at a cell
            try:
                testSheet_val = testSheet.iloc[rowNo, colNo]
            except:
                testSheet_val = np.nan

            # Comparing the values
            if (str(template_val) != str(testSheet_val)):
                cell = xl_rowcol_to_cell(rowNo, colNo)
                dfTemp = pd.DataFrame([[cell, template_val, testSheet_val]],
                                      columns=['Cell_Location', 'file1', 'file2'])
                # df = df.append(dfTemp)
                df = pd.concat([df, dfTemp])
    f = Frame(top)
    f.grid(row=4,column=2)
    #df = TableModel.getSampleData()
    top.table = pt = Table(f, dataframe=df,
                            showtoolbar=True, showstatusbar=True)
    pt.show()
   # print(df)

def setFile1(text):
    textFile1.delete(0,END)
    textFile1.insert(0,text)
    sortfile(text)
    return

def setFile2(text):
    textFile2.delete(0,END)
    textFile2.insert(0,text)
    sortfile(text)

    return

top = tki.Tk()
top.title("Excell Comperator")
top.geometry('800x600')




labelFile1 = Label(top, text="File1")
labelFile1.grid(padx=20,pady=20,row=1,column=1)
textFile1 = tki.Entry(top, width=40)
textFile1.grid(row=1,column=2)

buttonCompare = tki.Button(top, text="Choose File",command=lambda:(setFile1(filedialog.askopenfilename())))
buttonCompare.grid(pady=20,row=1,column=3);

labelFile2 = Label(top, text="File2")
labelFile2.grid(padx=20,pady=20,row=2,column=1)
textFile2 = tki.Entry(top, width=40)
textFile2.grid(row=2,column=2)

buttonCompare = tki.Button(top, text="Choose File",command=lambda:(setFile2(filedialog.askopenfilename())))
buttonCompare.grid(pady=20,row=2,column=3);




buttonCompare = tki.Button(top, text="Compare",command=lambda:comparefile())
buttonCompare.grid(pady=40,row=3,column=1);


top.mainloop()



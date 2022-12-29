#importing my modules
import tkinter
from tkinter import ttk
import openpyxl

#defining my function to load my exelsheet to my window
def exel_loader():
    path='C:\\Users\\Michael\\Desktop\\exel_viewer.xlsx'
    my_workbook=openpyxl.load_workbook(path)#loading the workbook into my python GUI
    sheet=my_workbook.active  #loading the sheet from my workbook into my python GUI
    
#listing values from my saved exel file
    lst_values=list(sheet.values)#get the data from my worksheet
    cls=lst_values[0]#populates column names to first row in the sheet
    
#populating data from exel file to this window
    tree=ttk.Treeview(window,column=cls ,show='headings')#creates columns for my exel data
    for cls_name in cls:#to populate text in column names
        tree.heading(cls_name,text=cls_name)

#populating other values in the rest of the rows
    for value_tuple in lst_values[1:]:
        tree.insert('',tkinter.END,value=value_tuple)
    tree.pack(expand=True,fill='both')#'fill' means content to fill both X and Y axises and
    #'expand' fills the entire sheet in the entire window

#creating a tkinter window to display my data
window=tkinter.Tk()#creating my main/root window
window.title('Exel_viewer')#naming my window
 

exel_loader()
window.mainloop()
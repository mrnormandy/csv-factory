# -*- coding: utf-8 -*-
"""
Created on Wed Sep 11 11:11:13 2019
@author: Gino Doran
CSV Creation and Apending Script 
"""
#========================================================= IMPORT LIBRARIES =============================================================================
import pandas as pd
import os
from os import listdir
import fnmatch
import tkinter as tk
from tkinter import messagebox

#========================================================= DEFINE CUSTOM EXCEPTION CLASS==================================================================
class Error(Exception):
    """Base class for exceptions in this module."""
    pass

class NoFilesFoundError(Error):
    """Exception raised for when there are no files found in a folder.

    Attributes:
        filepath -- input expression in which the error occurred
        message -- explanation of the error
    """

    def __init__(self, message, filepath):
        self.message = message
        self.filepath = filepath
#========================================================= DEFINE FUNCTIONS =============================================================================
def funcCSVOneFolder():
    #Append all CSV's in one folder to 1 CSV /Radionumber (1)
    csv_list = []
    os.chdir(mainfolderpath)
    print("Currently processing: " + mainfolderpath)
    print("Total CSV files found in folder= " + str(len(fnmatch.filter(os.listdir(mainfolderpath), '*.csv'))))
    if len(fnmatch.filter(os.listdir(mainfolderpath), '*.csv')) < 1: raise NoFilesFoundError(mainfolderpath,"No files found")
    if not os.path.exists(mainfolderpath + "\\CSV_OUTPUT"): #Create output folder
         os.makedirs(mainfolderpath + "\\CSV_OUTPUT")
    csv_list = [f for f in listdir(mainfolderpath) if f.endswith('.csv')]
    exportcsv = pd.concat(map(pd.read_csv, csv_list))
    exportcsv.to_csv(mainfolderpath + "\\" + "CSV_OUTPUT" + "\\DATA_" + os.path.basename(mainfolderpath) + ".csv", index=False, encoding='utf-8-sig') 

def funcCSVSubfolders():
    #Append all CSV's in each subfolders to 1 CSV /Radionumber (4)
    csv_list = []
    os.chdir(mainfolderpath)
    if not os.path.exists(mainfolderpath + "\\CSV_OUTPUT"): #Create output folder
         os.makedirs(mainfolderpath + "\\CSV_OUTPUT")
    for directoryname,subdirectorynames, filesnames in os.walk(mainfolderpath):
        for subdirectoryname in subdirectorynames:
            if subdirectoryname == "CSV_OUTPUT": continue
            print("Currently processing: " + mainfolderpath + "\\" + subdirectoryname)
            print("Total CSV files found in folder: '" + subdirectoryname + "'= " + str(len(fnmatch.filter(os.listdir(mainfolderpath+ "\\" + subdirectoryname), '*.csv'))))
            if len(fnmatch.filter(os.listdir(mainfolderpath+ "\\" + subdirectoryname), '*.csv')) < 1: continue
            os.chdir(mainfolderpath + "\\" + subdirectoryname)
            csv_list = [f for f in listdir(mainfolderpath + "\\" + subdirectoryname) if f.endswith('.csv')]
            exportcsv = pd.concat(map(pd.read_csv, csv_list))
            exportcsv.to_csv(mainfolderpath + "\\" + "CSV_OUTPUT" + "\\DATA_" + subdirectoryname + ".csv", index=False, encoding='utf-8-sig') 
        break

def funcExcelIndivdualOutput():
    #Create individual CSV files for each Excel files in 1 folder  /Radionumber (2)
    excel_list = []
    print("Currently processing: " + mainfolderpath)
    print("Total EXCEL files found in folder= " + str(len(fnmatch.filter(os.listdir(mainfolderpath), '*.xls*'))))
    if len(fnmatch.filter(os.listdir(mainfolderpath), '*.xls*')) < 1: raise NoFilesFoundError(mainfolderpath,"No files found")
    if not os.path.exists(mainfolderpath + "\\CSV_OUTPUT"): #Create output folder
         os.makedirs(mainfolderpath + "\\CSV_OUTPUT")
    os.chdir(mainfolderpath)
    excel_list = [f for f in listdir(mainfolderpath) if f.endswith('.xlsx') or f.endswith('.xlsm') ]
    for filename in excel_list:
        print('Converting: ' + filename)
        exportexcel = pd.read_excel(filename, sheetname, index_col =None)
        exportexcel.to_csv(mainfolderpath + "\\" + "CSV_OUTPUT" + "\\" + os.path.splitext(filename)[0] + ".csv", index=False, encoding='utf-8-sig') 

def funcExcelIndividualOutputSubfolders():
    #Convert all excels in a subfolder into individual csv outputs /Radionumber (5) 
    excel_list = []
    for directoryname,subdirectorynames, filesnames in os.walk(mainfolderpath):
            for subdirectoryname in subdirectorynames:
                if subdirectoryname == "CSV_OUTPUT": continue
                print("Currently processing: " + mainfolderpath + "\\" + subdirectoryname)
                print("Total EXCEL files found in folder: '" + subdirectoryname + "'= " + str(len(fnmatch.filter(os.listdir(mainfolderpath+ "\\" + subdirectoryname), '*.xls*'))))
                if len(fnmatch.filter(os.listdir(mainfolderpath+ "\\" + subdirectoryname), '*.xls*')) < 1: continue
                if not os.path.exists(mainfolderpath + "\\CSV_OUTPUT\\" + subdirectoryname):
                    os.makedirs(mainfolderpath + "\\CSV_OUTPUT\\" + subdirectoryname)
                os.chdir(mainfolderpath + "\\" + subdirectoryname)
                excel_list = [f for f in listdir(mainfolderpath + "\\" + subdirectoryname) if f.endswith('.xlsx') or f.endswith('.xlsm') ]
                for filename in excel_list:
                    print('Converting: ' + filename)
                    exportexcel = pd.read_excel(filename, sheetname, index_col =None)
                    exportexcel.to_csv(mainfolderpath + "\\" + "CSV_OUTPUT" + "\\" + subdirectoryname + "\\" + os.path.splitext(filename)[0] + ".csv", index=False, encoding='utf-8-sig')
            break

def funcExcelAppendedIndividualOutput():
    #Convert all excels in one folder to 1 csv output /Radionumber (3)
    excel_list = []
    excel_df = pd.DataFrame()
    print("Currently processing: " + mainfolderpath)
    print("Total EXCEL files found in folder: " + str(len(fnmatch.filter(os.listdir(mainfolderpath), '*.xls*'))))
    if len(fnmatch.filter(os.listdir(mainfolderpath), '*.xls*')) < 1: raise NoFilesFoundError(mainfolderpath,"No files found")
    os.chdir(mainfolderpath)
    if not os.path.exists(mainfolderpath + "\\CSV_OUTPUT"): #Create output folder
         os.makedirs(mainfolderpath + "\\CSV_OUTPUT")
    excel_list = [f for f in listdir(mainfolderpath) if f.endswith('.xlsx') or f.endswith('.xlsm') ]
    for filename in excel_list:
        print('Converting: ' + filename)
        exportexcel = pd.read_excel(filename, sheetname, index_col =None) 
        excel_df = excel_df.append(exportexcel)
    excel_df.to_csv(mainfolderpath + "\\" + "CSV_OUTPUT" + "\\" + "DATA_" + os.path.basename(mainfolderpath) + ".csv", index=False, encoding='utf-8-sig')

def funcExcelAppendedOutputSubfolders():
    #Convert all excels in subfolders to 1 csv output (6)
    excel_list = []
    excel_df = pd.DataFrame()
    for directoryname,subdirectorynames, filesnames in os.walk(mainfolderpath):
            for subdirectoryname in subdirectorynames:
                if subdirectoryname == "CSV_OUTPUT": continue
                print("Currently processing: " + mainfolderpath + "\\" + subdirectoryname)
                print("Total EXCEL files found in folder: '" + subdirectoryname + "'= " + str(len(fnmatch.filter(os.listdir(mainfolderpath+ "\\" + subdirectoryname), '*.xls*'))))
                if len(fnmatch.filter(os.listdir(mainfolderpath+ "\\" + subdirectoryname), '*.xls*')) < 1: continue
                if not os.path.exists(mainfolderpath + "\\CSV_OUTPUT\\" + subdirectoryname):
                    os.makedirs(mainfolderpath + "\\CSV_OUTPUT\\" + subdirectoryname)
                os.chdir(mainfolderpath + "\\" + subdirectoryname)
                excel_list = [f for f in listdir(mainfolderpath + "\\" + subdirectoryname) if f.endswith('.xlsx') or f.endswith('.xlsm') ]
                for filename in excel_list:
                    print('Converting: ' + filename)
                    exportexcel = pd.read_excel(filename, sheetname, index_col =None) 
                    excel_df = excel_df.append(exportexcel)
                    
                print("Appending file...")
                excel_df.to_csv(mainfolderpath + "\\" + "CSV_OUTPUT" + "\\" + subdirectoryname + "\\" + "DATA_" + subdirectoryname + ".csv", index=False, encoding='utf-8-sig')
            break
#========================================================= USER INPUT WITH TK GUI  =============================================================================
### Define Tk GUI functions
def click():
    input1 = str(txtFolderPath.get()).strip()
    txtOutput.delete(0.0, tk.END)
    txtOutput.insert(tk.END,input1)
    global mainfolderpath 
    mainfolderpath = input1
def exit_button():
    global selectedvalue
    selectedvalue = 0
    mainwindow.destroy()
def submit_button():
    input1 = str(txtFolderPath.get()).strip()
    mainfolderpath = input1
    if mainfolderpath == '':
        print("No folder path inputed")
        messagewindow = tk.Tk()
        messagewindow.withdraw()
        # message box display
        messagebox.showerror("Error", "No filepath was entered! Please verify.")
        messagewindow.destroy()
        return
    global selectedvalue
    selectedvalue = int(intSelectedValue.get())
    if selectedvalue == 0:
        print("No selection made")
        messagewindow = tk.Tk()
        messagewindow.withdraw()
        # message box display
        messagebox.showerror("Error", "No operation was selected! Please verify.")
        messagewindow.destroy()
        return
    global sheetname
    sheetname = str(txtSheetName.get()).strip()
    sheetname = "DATA" if sheetname == '' else sheetname   
    mainwindow.destroy()

#### Main Window
mainwindow = tk.Tk()
mainwindow.title("CSV FACTORY MAIN WINDOW")
mainwindow.configure(background="grey")
mainwindow.geometry('1350x450')

#Title label
lblTitle = tk.Label
lblTitle(mainwindow, text = "Please enter the path to the folder or folders, then click on 'SUBMIT FOLDERPATH' :", bg = "grey", fg="White", font ="none 12 bold").grid(row=1,column=0,sticky = "nw") #.place(x=0, y=0, anchor='nw')
#TextBox for folder path(s)
txtFolderPath = tk.Entry(mainwindow, width = 90, bg="white")
txtFolderPath.grid(row=1,column=1)
#Buttom submit filepath
btnSubmitFolderpath = tk.Button(mainwindow, text = "SUBMIT FOLDERPATH", width = 24, command = click,bg = "white", fg="grey", font ="none 12 bold") .grid(row=3, column =0,sticky = "n")
#label confirm filepath
lblConfirm = tk.Label
lblConfirm(mainwindow, text = "Confirm if filepath is correct:", bg = "grey", fg="White", font ="none 12 bold") .grid(row=4, column=0,sticky = "nw")
#Textbox outout box
txtOutput = tk.Text(mainwindow, width = 90, height = 2, bg="white", wrap = "word")
txtOutput.grid(row=5,column=0)
#visual buffer
lblBuffer = tk.Label
lblBuffer(mainwindow, text = "Perform operations on one folder only:", bg = "grey", fg="White", font ="none 12 bold") .grid(row=6, column=0,sticky = "nw")
#setup radio buttons and radiobutton values
intSelectedValue =tk.IntVar()
rbFunctionSelect = tk.Radiobutton

rbFunctionSelect(mainwindow,text = "(1) CSV: Append all CSV's in selected folder into 1 CSV", variable = intSelectedValue, value =1, indicatoron=1, font ="none 11",anchor=tk.W,relief = "sunken",width = 58).grid(row=7,column=0,sticky = tk.NW)
rbFunctionSelect(mainwindow,text = "(2) EXCEL: Create individual CSV files for each EXCEL files in selected folder", variable = intSelectedValue, value =2,indicatoron=1, font ="none 11",anchor=tk.W,relief = "sunken",width = 58).grid(row=8,column=0,sticky = tk.NW)
rbFunctionSelect(mainwindow,text = "(3) EXCEL: Convert all EXCEL files in selected folder to 1 CSV file", variable = intSelectedValue, value =3,indicatoron=1, font ="none 11",anchor=tk.W,relief = "sunken",width = 58).grid(row=9,column=0,sticky = tk.NW)
#visual buffer
lblBuffer = tk.Label
lblBuffer(mainwindow, text = "Perform operations on multiple subfolders:", bg = "grey", fg="White", font ="none 12 bold") .grid(row=10, column=0,sticky = "nw")

rbFunctionSelect(mainwindow,text = "(4) CSV: Append all CSV's in all subfolders into 1 CSV file per subfolder", variable = intSelectedValue, value =4,indicatoron=1, font ="none 11",anchor=tk.W,relief = "sunken",width = 58).grid(row=11,column=0,sticky = tk.NW)
rbFunctionSelect(mainwindow,text = "(5) EXCEL Convert all EXCEL files in all subfolders into 1 CSV file per subfolder ", variable = intSelectedValue, value =5,indicatoron=1, font ="none 11",anchor=tk.W,relief = "sunken",width = 58).grid(row=12,column=0,sticky = tk.NW)
rbFunctionSelect(mainwindow,text = "(6) EXCEL: Convert all EXCEL files in all subfolders into 1 CSV file per subfolder", variable = intSelectedValue, value =6,indicatoron=1, font ="none 11",anchor=tk.W,relief = "sunken",width = 58).grid(row=13,column=0,sticky = tk.NW)
#visual buffer
lblBuffer = tk.Label
lblBuffer(mainwindow, text = "", bg = "grey", fg="White", font ="none 12 bold") .grid(row=14, column=0,sticky = "nw")
#Sheetname label
lblTitle = tk.Label
lblTitle(mainwindow, text = "(IF converting EXCEL files, please input the SHEET NAME. If the SHEET is already:'DATA' leave this blank)", bg = "grey", fg="White", font ="none 11 italic").grid(row=15,column=0,sticky = "nw") #.place(x=0, y=0, anchor='nw')
#TextBox for folder path(s)
txtSheetName = tk.Entry(mainwindow, width = 90, bg="white")
txtSheetName.grid(row=15,column=1)
#visual buffer
lblBuffer = tk.Label
lblBuffer(mainwindow, text = "", bg = "grey", fg="White", font ="none 12 bold") .grid(row=16, column=0,sticky = "nw")
#exit button
btnSubmit = tk.Button(mainwindow,text = "SUBMIT", width = 8, command = submit_button, font ="none 12 bold") .grid(row = 20, column =0,sticky = tk.NE)
btnCancel = tk.Button(mainwindow,text = "CANCEL", width = 8, command = exit_button, font ="none 12 bold") .grid(row = 20, column =0,sticky = tk.NW)
### WIndow loop
mainwindow.mainloop()
#========================================================= EXECUTE PROGRAM  =============================================================================
#If Statment to go through selected option
try:
    if selectedvalue == 1:
        funcCSVOneFolder()
    elif selectedvalue ==2:
        funcExcelIndivdualOutput()
    elif selectedvalue ==3:
        funcExcelAppendedIndividualOutput()
    elif selectedvalue ==4:
        funcCSVSubfolders()
    elif selectedvalue ==5:
        funcExcelIndividualOutputSubfolders()
    elif selectedvalue ==6:
        funcExcelAppendedOutputSubfolders()
except NoFilesFoundError:
    #Custom Exception 
    print("There were no files found in the folder. Please confirm nd try again")
    messagewindow = tk.Tk()
    messagewindow.withdraw()
    # message box display
    messagebox.showerror("Error", "There were no files found in the folder. Please confirm and try again.")
    messagewindow.destroy()
except FileNotFoundError:
    print("The inputed Folder path was not found or does not exist. Please confirm and try again")
    messagewindow = tk.Tk()
    messagewindow.withdraw()
    # message box display
    messagebox.showerror("Error", "The inputed Folder path was not found or does not exist. Please confirm and try again")
    messagewindow.destroy()
except Exception as error: # catches built in exceptions and ignores KeyboardInterrupt, SystemExit, and GeneratorExit
    print("Unknown Error!. Please try again")
    messagewindow = tk.Tk()
    messagewindow.withdraw()
    # message box display
    messagebox.showerror("Error", "Unknown Error!. Please try again. Description: " + str(error))
    messagewindow.destroy()
else:
    print("Program completed succesfully")
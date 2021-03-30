import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd

root = tk.Tk()

root.title("XLSX CONVERTER")
canvas1 = tk.Canvas(root, width=300, height=300, bg='red', relief='raised')
canvas1.pack()

label1 = tk.Label(root, text='CONVERSION TOOL', bg='gold')
label1.config(font=('helvetica', 20))
canvas1.create_window(150, 60, window=label1)


def getExcel():
    global read_file

    import_file_path = filedialog.askopenfilename()
    read_file = pd.read_excel(import_file_path,sheet_name="SRC_CMR_DATA")


browseButton_Excel = tk.Button(text="      Import Excel File     ", command=getExcel, bg='cyan2', fg='white',
                               font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 130, window=browseButton_Excel)


def convertToCSV():
    global read_file

    export_file_path = filedialog.asksaveasfilename(defaultextension='.csv')
    read_file.to_csv(export_file_path, index=None, header=True)


saveAsButton_CSV = tk.Button(text='Convert Excel to CSV', command=convertToCSV, bg='green3', fg='white',
                             font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 180, window=saveAsButton_CSV)

def convertToTXT():
    global read_file

    export_file_path = filedialog.asksaveasfilename(defaultextension='.txt')
    read_file.to_csv(export_file_path, index=None, header=True)

saveAsButton_TXT = tk.Button(text='Convert Excel to TXT', command=convertToTXT, bg='SpringGreen2', fg='white',
                             font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 230, window=saveAsButton_TXT)

def exitApplication():
    MsgBox = tk.messagebox.askquestion('Exit Application', 'Are you sure you want to exit the application',
                                       icon='warning')
    if MsgBox == 'yes':
        root.destroy()


exitButton = tk.Button(root, text='       Exit Application     ', command=exitApplication, bg='firebrick2', fg='white',
                       font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 280, window=exitButton)

root.mainloop()

from tkinter import filedialog
from tkinter import *
from turtle import width
from InvoiceGen import InvoiceGen
import os

fileJobCodes = 'job_codes.xlsx'

###### tkinter methods #####
root = Tk()
root.title('Indeed Extractor')
root.geometry('300x200')

def open():
  global fileNames
  root.fileNames = filedialog.askopenfilenames(initialdir="/", title="Select Files", filetypes=(('csv files','*.csv'),("all files","*.*")))

  label.config(text="")

  newLabel = ""

  tempFns = list(root.fileNames).sort(key=sortFileNamesByDateModified)

  print('fileNames', root.fileNames)
  print('sorted', tempFns)

  for fn in root.fileNames:
    # print(os.path.getmtime(fn))
    newLabel += fn + '\n'
  
  root.newLabel = newLabel
  
  label.config(text=newLabel)


def run():
  outputFileName = text.get("1.0",'end-1c')
  newInstance = InvoiceGen(root.fileNames, fileJobCodes, outputFileName)
  newInstance.run()

def sortFileNamesByDateModified(fn):
  return os.path.getmtime(fn)

def stringDatetoDateTime(stringDate):
  dateMap = {
    'Jan': 1,
    'Feb': 2,
    'Mar': 3,
    'Apr': 4,
    'May': 5,
    'Jun': 6,
    'Jul': 7,
    'Aug': 8,
    'Sep': 9,
    'Oct': 10,
    'Nov': 11,
    'Dec': 12
  }

  
runBtn = Button(root, text="Run", command=run)
runBtn.config(width = 20)
runBtn.pack(pady=10)

outputlabel = Label(root, text="Name output file")
outputlabel.pack()

text= Text(root, width= 20, height= 2, background=
"white",foreground="black")
text.insert(INSERT, "Indeed Invoice Breakdowns")
text.pack()

chooseFilesButton = Button(root, text="Choose CSV Files", command=open).pack(pady = 20)
label = Label(root, text="")
label.pack()

root.mainloop()
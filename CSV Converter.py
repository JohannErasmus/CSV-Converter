import pandas as pd
#import openpyxl
#import numpy
from xlsxwriter import *
from datetime import *
from tkinter import *
from tkinter import filedialog
from tkcalendar import *

root = Tk()
root.title('CSV Processor')
root.geometry('370x310')

#Return file name 
def get_file_name(file_entry):
    file_name = filedialog.askopenfilename(title='Select CSV to read', filetypes=(('ZIP files', '.zip'),('CSV files','*.csv'),))
    file_entry.delete(0,END)
    file_entry.insert(0,file_name)
    instructionLbl.config(text='Click Read Button')

#Read csv form location in storage, add Test Number column and format columns
def readClick():
    global df2
    df = pd.read_csv(entry_csv.get(),low_memory=False,compression='zip', header=None,names=('Date & Time','Process Steps',"CC Long Temp",'CC Short Temp','AL Temp','CA Temp','CA RPM','CA Pressure','CC Pressure','QA RPM','QA Airflow','QA Pressure','RV RPM','Gas kg/h','DC TT Temp','CYC1 TT Temp','CYC2 TT Temp','Product Pump RPM','Product flow l/h','ID RPM','DC PT Pressure'))
    df2 = df.loc[df['Process Steps'].isin(['1600', '2000', '2001', '2003', '3002'])]
    df2.insert(0, 'Test Number', '', allow_duplicates=True)
    df2['Date & Time'] = pd.to_datetime(df2['Date & Time'])
    df2[['Process Steps', 'Test Number']] = df2[['Process Steps', 'Test Number']].apply(pd.to_numeric)
    instructionLbl.config(text='Choose a Date')

#Grab event that gets the date from calendar input
def grabDate(event):
    global date1
    date1 = cal.get_date()
    date1 = datetime.strptime(date1,'%m/%d/%y').date()
    instructionLbl.config(text='Click Write Button')
    dateLbl.config(text='Date selected:' + str(date1))
                          
#Filter dataframe by date selected
def filterDate():
    global df3, df4, dm
    df3 = df2.copy()
    df3['Date & Time'] = pd.to_datetime(df3['Date & Time'], format='%Y-%m-%d')
    df3['DateTime'] = df3['Date & Time'].dt.date
    df4 = df3.loc[(df3['DateTime'] == date1)]
    df4 = df4.drop(axis='columns', columns='DateTime')
    df4.reset_index(inplace=True, drop=True)
    dm = df4[['Date & Time']].copy()
    dm['Day'] = dm['Date & Time'].dt.day
    dm['Month'] = dm['Date & Time'].dt.month
    dm['Year'] = dm['Date & Time'].dt.year

#Create Test number and get file name date
def testNumber(index, n): 
    global dateStr
    day_n = dm.loc[index, 'Day']
    month_n = dm.loc[index, 'Month']
    year_n = dm.loc[index, 'Year']
    dateStr = str(day_n) + '-' + str(month_n) + '-' + str(year_n)
    s = str(n).zfill(3)
    test_nr = str(day_n) + str(month_n) + s
    return int(test_nr)

#Calls filterDate function then adds the Test number, then sets the index, formats the columns to float and rounds
def convert():   
    global n, df5
    n = 1 
    switchVar = 0
    filterDate()
    df5 = df4.copy()
    for index, row in df5.iterrows():
        nr = testNumber(index, n)
        process_step = df5.loc[index, 'Process Steps']
        if process_step == 1600 and switchVar == 0: 
            df5.loc[index, ['Test Number']] = nr  
        elif process_step == 2000:
            df5.loc[index, ['Test Number']] = nr
        elif process_step == 2001:
            df5.loc[index, ['Test Number']] = nr
        elif process_step == 2003:
            df5.loc[index, ['Test Number']] = nr
        elif process_step == 3002:
            df5.loc[index, ['Test Number']] = nr
            switchVar = 1
        elif process_step == 1600 and switchVar == 1:
            n += 1
            nr = testNumber(index, n)
            df5.loc[index, ['Test Number']] = nr
            switchVar = 0   
    df5.set_index('Test Number', inplace=True, drop=True)
    df5[["CC Long Temp", 'CC Short Temp','AL Temp', 'CA Temp','CA RPM','CA Pressure','CC Pressure','QA RPM','QA Airflow', 'QA Pressure','RV RPM', 'Gas kg/h','DC TT Temp', 'CYC1 TT Temp','CYC2 TT Temp', 'Product Pump RPM','Product flow l/h', 'ID RPM','DC PT Pressure']] \
        = df5[["CC Long Temp", 'CC Short Temp','AL Temp', 'CA Temp','CA RPM','CA Pressure','CC Pressure','QA RPM','QA Airflow', 'QA Pressure','RV RPM', 'Gas kg/h','DC TT Temp', 'CYC1 TT Temp','CYC2 TT Temp', 'Product Pump RPM','Product flow l/h', 'ID RPM','DC PT Pressure']].apply(pd.to_numeric)
    df5 = df5.round({"CC Long Temp":0,'CC Short Temp':0,'AL Temp':0, 'CA Temp':0,'CA RPM':0,'CA Pressure':0,'CC Pressure':0,'QA RPM':0,'QA Airflow':0, 'QA Pressure':0,'RV RPM':0, 'Gas kg/h':2,'DC TT Temp':0, 'CYC1 TT Temp':0,'CYC2 TT Temp':0, 'Product Pump RPM':0,'Product flow l/h':1, 'ID RPM':0,'DC PT Pressure':0})

#Creates a filename, a writing engine with xlsxwriter, formats the headers and column size, and writes to a new xlsx file.
def createFile():
    global fileName
    fileName = 'PCE Test Log ' + dateStr + '.xlsx' 
    (max_row, max_col) = df5.shape
    writer = pd.ExcelWriter(fileName, engine='xlsxwriter')
    df5.to_excel(writer, sheet_name='Data',freeze_panes=(1,1))
    workbook = writer.book
    worksheet = writer.sheets['Data']
   
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'align': 'center',
        'fg_color': '#a9d1de',
        'border': 2})
    
    worksheet.write(0, 0, 'Test Number', header_format)
    for col_num, value in enumerate(df5.columns.values):
       worksheet.write(0, col_num + 1, value, header_format)
    worksheet.autofilter(0, 0, max_row, 0)
    worksheet.autofit()
    format_col_width(worksheet)
    writer.close()

#Column width control
def format_col_width(worksheet):
    worksheet.set_column(1, 1, 17)
    worksheet.set_column('C:C', 7.5)
    worksheet.set_column('D:H', 5.3)
    worksheet.set_column('I:J', 8)
    worksheet.set_column('K:K', 4.5)    
    worksheet.set_column('L:L', 7)
    worksheet.set_column('M:M', 8)
    worksheet.set_column('N:N', 4.5) 
    worksheet.set_column('O:P', 5.3)
    worksheet.set_column('Q:R', 7)
    worksheet.set_column('S:T', 10)
    worksheet.set_column('U:U', 4.3)
    worksheet.set_column('V:V', 8)

#Runs convert function on the dataframe, then runs createFile function
def writeClick():
    convert() 
    createFile()
    instructionLbl.config(text='Done')
   
    #// Future // with pd.ExcelWriter('Main PCE Test Log.xlsx',mode='a',engine='openpyxl',if_sheet_exists='overlay') as writer:
    #       df5.to_excel(writer, sheet_name='Data',header=None, startrow=writer.sheets['Data'].max_row, index= 'Test Number', freeze_panes=(1,1))

#Create GUI objects
instructionLbl = Label(root, text='Choose CSV Location', font=('Times New Roman', 18))
dateLbl = Label(root, text='No Date Selected', font=('Times New Roman', 18))
entry_csv = Entry(root, text='', width=50)
browseButton = Button(root, text='Browse', command=lambda:get_file_name(entry_csv))
readButton = Button(root, text='Read', command=readClick)
writeButton = Button(root, text='Write', command=writeClick)
cal = Calendar(root)
cal.bind('<<CalendarSelected>>', grabDate)

#Place objects on GUI
cal.grid(row=3, column=2)
instructionLbl.grid(row=0, column=2)
dateLbl.grid(row=4, column=2)
entry_csv.grid(row=1, column=2)
browseButton.grid(row=1, column=1)
readButton.grid(row=2, column=1)
writeButton.grid(row=3, column=1)

root.mainloop()
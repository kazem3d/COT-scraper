import xlrd
import sqlite3
import xlwt 
from xlwt import Workbook 
  



address='annual.xls'

#open xls file
excel_reader=xlrd.open_workbook(address) 
sheet = excel_reader.sheet_by_index(0) 
sheet.cell_value(0,0) 

#create and connect to report_datebase

conn=sqlite3.connect('report_database')
curser=conn.cursor()

#create main table in database   
curser.execute('''

    CREATE TABLE IF NOT EXISTS main (

    id integer PRIMARY KEY,
    currency TEXT NOT NULL ,
    date TEXT  ,
    long TEXT  ,
    short TEXT,
    long_change TEXT,
    short_change TEXT,
    long_change_percent TEXT,
    short_change_percent TEXT,
    net_position TEXT

    );''')

#function for writing a tupel in main table
def record_in_database(data_tupel):

    curser.execute('INSERT INTO main (currency,date,long,short,long_change,short_change,long_change_percent,short_change_percent,net_position)  VALUES (?,?,?,?,?,?,?,?,?)' ,data_tupel)
    conn.commit()

currency_list=[
    'JAPANESE YEN - CHICAGO MERCANTILE EXCHANGE',
    'SWISS FRANC - CHICAGO MERCANTILE EXCHANGE',
    'CANADIAN DOLLAR - CHICAGO MERCANTILE EXCHANGE',
    'BRITISH POUND STERLING - CHICAGO MERCANTILE EXCHANGE',
    'U.S. DOLLAR INDEX - ICE FUTURES U.S.',
    'EURO FX - CHICAGO MERCANTILE EXCHANGE',
    'NEW ZEALAND DOLLAR - CHICAGO MERCANTILE EXCHANGE',
    'AUSTRALIAN DOLLAR - CHICAGO MERCANTILE EXCHANGE',
]

sheet_name=['JPY','CHF','CAD','GBP','USD','EUR','NZD','AUD']
# Workbook is created 
wb = Workbook() 
  
# add_sheet is used to create sheet. 
sheet_name_obj=[]
for i in sheet_name:
    sheet_name_obj.append(wb.add_sheet(i))

sheet_header_list=['name','date','long','short','change_long','change_short','long_change_percent','short_change_percent','net_position']


for currency_id, currency_name in enumerate(currency_list) :
    row_id=0
    #writing header
    for cell,cell_data in enumerate(sheet_header_list) :
                sheet_name_obj[currency_id].write(0, cell,cell_data) 
                # print(0, cell,cell_data) 

    for i in range(0,(sheet.nrows)):
        row=0
        row=sheet.row_values(i)

        if row[0] == currency_name:
            print(row[0],row[2],row[8],row[9],row[38],row[39])
            long_change_percent=(row[38]/(row[8]-row[38]))
            short_change_percent=(row[39]/(row[9]-row[39]))
            net_position=row[8]-row[9]
            row_tuple=(row[0],row[2],row[8],row[9],row[38],row[39],long_change_percent,short_change_percent,net_position,)
            record_in_database(row_tuple)
            for cell,cell_data in enumerate(row_tuple) :
                sheet_name_obj[currency_id].write(row_id+1, cell,cell_data) 
                    
            row_id+=1

wb.save('COT report.xls') 
            
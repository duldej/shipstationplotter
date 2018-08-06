import os
import pandas as pd
import csv
import openpyxl
import xlsxwriter

#from datetime import datetime
#import matplotlib.pyplot as pyplot
import numpy as np
#import array
arr=np.array([])
#set working directory
os.chdir("/Users/Jason/Desktop/")
table=0
tablet1=0
n=0
dates=[]
items=[]
quantity=[]
table=[]

#get filename input
n=input("please enter top n results:")
x=input("Please enter csv filename:")
n=int(n)
#create a list
#with open(x) as csv_file:
df=pd.read_csv(x)
    
    #csv.reader(csv_file, delimiter=",")
 
#convert dates to usable date code

pd.to_datetime(df['OrderDate']).apply(lambda x: x.date())
df['Month'] = pd.DatetimeIndex(df['OrderDate']).month
df2=df[['Month','Description','Quantity']].copy()
maxmonths=max(df['Month'])
#drop rows with missing data
df2.dropna(inplace=True)
    

#create pivot table
table=df2.pivot_table(index='Month', columns='Description',values='Quantity',aggfunc='sum')

#create totals
table.loc['sum'] = table.sum()


#transpose matrix
tablet=table.T

#take top n values
tablet1 = tablet.sort_values('sum',ascending = False).head(n)

#eliminate the blanks in the table
tablet1=tablet1.fillna(0)
#convert pivot table to df
flattened = pd.DataFrame(tablet1.to_records())

# Create a Pandas Excel writer using XlsxWriter as the engine.
sheet_name = 'Sheet1'
writer     = pd.ExcelWriter('pandas_chart_line.xlsx', engine='xlsxwriter')
flattened.to_excel(writer, sheet_name=sheet_name)

# Access the XlsxWriter workbook and worksheet objects from the dataframe.
workbook  = writer.book
worksheet = writer.sheets[sheet_name]

# Create a chart object.
chart = workbook.add_chart({'type': 'line'})

# Configure the series of the chart from the dataframe data.
for i in range(0,n):
    col = i + 1
    chart.add_series({
        'name':       ['Sheet1', i+1,1],
        'categories': ['Sheet1',0,2,0,maxmonths],
        'values':     ['Sheet1',i+1,2,i+1,maxmonths ],
    })

# Configure the chart axes.
chart.set_x_axis({'name': 'Months'})
chart.set_y_axis({'name': 'Units Sold', 'major_gridlines': {'visible': False}})

# Insert the chart into the worksheet.
worksheet.insert_chart('K2', chart)

# Close the Pandas Excel writer and output the Excel file.
writer.save()















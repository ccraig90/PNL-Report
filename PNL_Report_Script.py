import win32com.client
import os
import datetime as datetime
import numpy as np
import pandas as pd
import xlsxwriter
import datetime
import io

def subject_name(file_text):
    now = datetime.datetime.now()
    weekday = now.weekday()
    if weekday == 0:
        today = now - datetime.timedelta(days=3)
        yesterday = today - datetime.timedelta(days=1)
    elif weekday == 1:
        today = now - datetime.timedelta(days=1)
        yesterday = today - datetime.timedelta(days=3)
    else:
        today = now - datetime.timedelta(days=1)
        yesterday = today - datetime.timedelta(days=1)
    current = today.strftime("%Y-%m-%d")
    current = str(file_text)+str(current)
    PNL_Report_Date = today.strftime("%m.%d.%Y")
    PNL_Report_Date = str(PNL_Report_Date)+'PNL Discrepancy'
    now = datetime.datetime.now()
    PNL_Report_Write_to_Date = now.strftime('%m.%d.%Y')
    PNL_Report_Write_to_Date = str(PNL_Report_Write_to_Date)+'PNL Discrepancy'
    today = file_text + str(today)
    yesterday = yesterday.strftime("%Y-%m-%d")
    yesterday = file_text + str(yesterday)
    return current,yesterday,PNL_Report_Date,PNL_Report_Write_to_Date

def correct_position_type(Inventory):
    x = Inventory['Position']
    x = pd.to_numeric(x)
    Inventory['Position'] = x
    x = Inventory['P&L']
    x = pd.to_numeric(x)
    Inventory['P&L'] = x
    return Inventory


file_text = 'Inventory Margin Report for '
x = subject_name(file_text)

"""
Pull and Generate new QTY_DSP_Cleared_Positions file
"""
PNL_Report_Date = x[2]
PNL_Report_File_Most_Recent = 'P:/2. Corps/PNL_Daily_Report/Reports/PNL_Report.xlsx'
Cleared_Yesterday_PNL_Report = pd.read_excel(PNL_Report_File_Most_Recent,sheet_name = 'HT QTY DSP')
Cleared_Yesterday_PNL_Report = Cleared_Yesterday_PNL_Report[['Security','Account','Cusip','QTY DSP','Position Notes']]
Cleared_Yesterday_PNL_Report.dropna(inplace = True)
QTY_DSP_Cleared_Positions = 'P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/QTY_DSP_Cleared_Positions.xlsx'
QTY_DSP_Cleared_Positions= pd.read_excel(QTY_DSP_Cleared_Positions,index = False)
QTY_DSP_Cleared_Positions = QTY_DSP_Cleared_Positions.append(Cleared_Yesterday_PNL_Report)

QTY_DSP_Cleared_Positions.drop_duplicates(keep='first',inplace = True)
writer = pd.ExcelWriter('P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/QTY_DSP_Cleared_Positions.xlsx', engine='xlsxwriter')
QTY_DSP_Cleared_Positions.to_excel(writer)
writer.save()


today = x[0]
yesterday = x[1]
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()

i = 0
while i < 20:
    if message.Subject == today:
        try:
            attachments = message.Attachments
            attachment = attachments.Item(1)
            attachment.SaveASFile('P:/2. Corps/PNL_Daily_Report/HT_Files/' + str(today)+'.xlsx') #Saves to the attachment to current folder
            print('HT File Found')
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()


file_text = 'Report "TW 16 22" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_16_22 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22.transpose()
Bloomberg_Inventory_16_22.reset_index(inplace = True)
Bloomberg_Inventory_16_22.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_16_22.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22.transpose()
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[:-1]
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[2:]
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[0].str.split(',',expand=True)
Bloomberg_Inventory_16_22.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_16_22.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_16_22 = correct_position_type(Bloomberg_Inventory_16_22)

file_text = 'Report "TW 1 5" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:

    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_1_5 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5.transpose()
Bloomberg_Inventory_1_5.reset_index(inplace = True)
Bloomberg_Inventory_1_5.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_1_5.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5.transpose()
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[:-1]
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[2:]
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[0].str.split(',',expand=True)
Bloomberg_Inventory_1_5.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_1_5.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_1_5 = correct_position_type(Bloomberg_Inventory_1_5)

file_text = 'Report "TW 6 10" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_6_10 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10.transpose()
Bloomberg_Inventory_6_10.reset_index(inplace = True)
Bloomberg_Inventory_6_10.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_6_10.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10.transpose()
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[:-1]
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[2:]
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[0].str.split(',',expand=True)
Bloomberg_Inventory_6_10.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_6_10.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_6_10 = correct_position_type(Bloomberg_Inventory_6_10)

file_text = 'Report "TW 11 15" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_11_15 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15.transpose()
Bloomberg_Inventory_11_15.reset_index(inplace = True)
Bloomberg_Inventory_11_15.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_11_15.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15.transpose()
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[:-1]
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[2:]
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[0].str.split(',',expand=True)
Bloomberg_Inventory_11_15.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_11_15.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_11_15 = correct_position_type(Bloomberg_Inventory_11_15)

Bloomberg_Inventory = pd.concat([Bloomberg_Inventory_16_22,
                                 Bloomberg_Inventory_1_5,
                                 Bloomberg_Inventory_6_10,
                                 Bloomberg_Inventory_11_15],ignore_index = True)

"""
Read in Excel files from Hilltop and Bloomberg

"""
# Bloomberg_Inventory = pd.read_excel('C:/Users/ccraig/Desktop/PNL Project/8.21.2019 Inventory TW.xlsx')
# x = Bloomberg_Inventory[4:]
# x.rename(columns={x.columns[0]: 'Security',x.columns[1]:'P&L',x.columns[2]:'Cusip',x.columns[3]:'Position',x.columns[4]:'Symbol'}, inplace=True)
# Bloomberg_Inventory.rename(columns={'CUSIP': 'Cusip'}, inplace=True)
Bloomberg_Inventory['Cusip'] = Bloomberg_Inventory['Symbol'].str[1:10]
Bloomberg_Inventory = Bloomberg_Inventory[['Cusip', 'P&L', 'Security', 'Position','Symbol','Book']]
Bloomberg_Inventory['Position'] = Bloomberg_Inventory['Position']*1000
# Bloomberg_Inventory['Cusip'] = Bloomberg_Inventory['Cusip'].astype(str)
print(today)
print(yesterday)
Recent = 'C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\' + str(today)+'.xlsx'#Recent = 'C:/Users/ccraig/Desktop/PNL Project/'+str(today)+'.xlsx'
Old = 'C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\' + str(yesterday)+'.xlsx'
Hilltop_Recent_x = pd.read_excel(io=Recent, sheet_name='Detail')
# Hilltop_Recent_x['Cusip'] = Hilltop_Recent['Cusip'].astype(str)
Hilltop_Old_y = pd.read_excel(io=Old, sheet_name='Detail')
# Hilltop_Old_y['Cusip'] = Hilltop_Old_y['Cusip'].astype(str)
Hilltop_Recent_s = pd.read_excel(io=Recent, sheet_name='Summary')
Hilltop_Old_s = pd.read_excel(io=Old, sheet_name='Summary')
Hilltop_Recent_s = Hilltop_Recent_s.head(10)
Hilltop_Old_s = Hilltop_Old_s.head(10)
Hilltop_Recent_x['Cusip_group_by'] = Hilltop_Recent_x['Cusip']
Hilltop_Recent_x['Cusip_group_by'] = 'C'+ Hilltop_Recent_x['Cusip_group_by']
Bloomberg_Inventory = Bloomberg_Inventory.groupby(['Cusip']).agg({'P&L':'sum',
                                                                   'Security':'first',
                                                                   'Position':'sum',
                                                                   'Symbol':'first',
                                                                 'Book':'first'})


Hilltop_Recent = Hilltop_Recent_x.groupby(['Cusip_group_by']).agg({'Quantity':'sum',
                                                                   'Unreal PNL':'sum',
                                                                   'Real PNL':'sum',
                                                                   'Requirement':'sum',
                                                                   'Cusip':'first',
                                                                   'Description':'first',
                                                                   'Price':'mean'})

Hilltop_Old_y['Cusip_group_by'] = Hilltop_Old_y['Cusip']
Hilltop_Old = Hilltop_Old_y.groupby(['Cusip_group_by']).agg({'Quantity':'sum',
                                                             'Unreal PNL':'sum',
                                                             'Real PNL':'sum',
                                                             'Requirement':'sum',
                                                             'Cusip':'first',
                                                             'Description':'first',
                                                             'Price':'mean'})

"""
Merge Hilltop Recent and Old together

"""

Hilltop_Recent = pd.merge(Hilltop_Recent, Hilltop_Old, on='Cusip', how='left')

Hilltop_Recent = pd.merge(Hilltop_Recent, Bloomberg_Inventory, on='Cusip', how='outer')

Hilltop_Recent_x = Hilltop_Recent_x[['Cusip','Account Name']]

Hilltop_Recent = pd.merge(Hilltop_Recent, Hilltop_Recent_x,on='Cusip',how='outer')

Hilltop_Recent = Hilltop_Recent.fillna(0)
Hilltop_Recent.loc[Hilltop_Recent['Security']==0,'Security'] = Hilltop_Recent['Description_x']

"""
Calculate nessicary values

"""

Hilltop_Recent['Real_PNL_Change'] = Hilltop_Recent['Real PNL_x'] - Hilltop_Recent['Real PNL_y']
Hilltop_Recent['Real Discrepancy'] = Hilltop_Recent['Real_PNL_Change'] - Hilltop_Recent['P&L']
Hilltop_Recent['Quantity Change'] = Hilltop_Recent['Quantity_x'] - Hilltop_Recent['Quantity_y']

Hilltop_Individual_Book_Summary = Hilltop_Recent

Hilltop_Recent.rename(columns={'Quantity Change': 'HT Quantity Change',
                               'Quantity_x': 'HT New Quantity',
                               'Quantity_y': 'HT Old Quantity',
                               'Quantity_x_y': '-  HT Quantity  =',
                               'Real Discrepancy_y': 'TW vs. HT Real Discrepancy',
                               'Position': 'TW Quantity',
                               'Real PNL_x': 'HT New Real PNL',
                               'Account Name': 'Account',
                               'Quantity_y':'HT Old Quantity',
                               'Unreal PNL_x':'HT New Unreal PNL',
                               'Unreal PNL_y':'HT Old Unreal PNL',
                               'Real PNL_y':'HT Old Real PNL',
                               'P&L':'TW PNL',
                               'Quantity_y':'HT Old Quantity',
                               'Price_x':'Price'}, inplace=True)

# # Hilltop_Recent['TW Quantity'] = Hilltop_Recent['TW Quantity'] * 1000

Hilltop_Recent['HT Change in Quantity'] = Hilltop_Recent['HT New Quantity']-Hilltop_Recent['HT Old Quantity']
Hilltop_Recent['TW - HT Quantity Discrepancy'] = Hilltop_Recent['TW Quantity']-Hilltop_Recent['HT New Quantity']
Hilltop_Recent['Real PNL Change'] = Hilltop_Recent['HT New Real PNL']-Hilltop_Recent['HT Old Real PNL']
Hilltop_Recent['HT Real PNL Change'] = Hilltop_Recent['HT New Real PNL']-Hilltop_Recent['HT Old Real PNL']
Hilltop_Recent['Adj Unreal PNL Change'] = Hilltop_Recent['HT New Unreal PNL']-Hilltop_Recent['HT Old Unreal PNL']+Hilltop_Recent['HT Real PNL Change']
Hilltop_Recent['HT-TW PNL Discrepancy'] = Hilltop_Recent['HT Real PNL Change']-Hilltop_Recent['TW PNL']
Hilltop_Recent['Requirement Change'] = Hilltop_Recent['Requirement_x']-Hilltop_Recent['Requirement_y']
Hilltop_Recent['Filter Column'] = Hilltop_Recent['TW - HT Quantity Discrepancy'] + Hilltop_Recent['HT Change in Quantity'] + Hilltop_Recent['Adj Unreal PNL Change'] + Hilltop_Recent['HT-TW PNL Discrepancy'] + Hilltop_Recent['Requirement Change']
Hilltop_Recent = Hilltop_Recent[(Hilltop_Recent['Filter Column'] != 0)]

Hilltop_Recent = Hilltop_Recent[[
                                 'HT Quantity Change',                         
                                 'Security',                                       #A
                                 'Cusip',                                          #B
                                 'Account',                                        #C
                                 'Price',                                          #D
                                 'TW Quantity',                                    #E
                                 'HT New Quantity',                                #F
                                 'TW - HT Quantity Discrepancy',                   #G
                                 'HT Change in Quantity',                          #H
                                 'HT New Unreal PNL',                              #I
                                 'HT Old Unreal PNL',                              #J
                                 'Real PNL Change',                                #K
                                 'Adj Unreal PNL Change',                          #L
                                 'TW PNL',                                         #M
                                 'HT-TW PNL Discrepancy',                          #N
                                 'Requirement Change']]#9    

"""
# Drop Duplicated Values for Cusip and HT Quantity Change

"""
Hilltop_Recent = Hilltop_Recent.drop_duplicates(['Cusip', 'HT Quantity Change'])


"""
# Set up excel file naming path w/ today's date

"""
time = datetime.datetime.today()
current = time.strftime("%m.%d.%Y")
"""
# Write file to excel

"""
Hilltop_Recent.drop(
    [
        'HT Quantity Change'
    ],
    axis=1, inplace=True)


Hilltop_Individual_Summary = Hilltop_Recent
Hilltop_Recent.sort_values('TW PNL', axis=0, ascending=False, inplace=True)

filepath = 'P:/2. Corps/PNL_Daily_Report/Reports/PNL_Report.xlsx'

writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
"""
            Summary Code

"""
Daily_Change_x = pd.read_excel(io=Recent, sheet_name='Summary')
Daily_Change_y = pd.read_excel(io=Old,sheet_name='Summary')
Daily_Change_x = Daily_Change_x[12:]
Daily_Change_y = Daily_Change_y[12:]
Daily_Change_x.rename(columns={'Account Number': 'Inventory Sub Totals',
                                'Total Available Funds':'Account Name',
                                'Unnamed: 2':'Position Type',
                                'Unnamed: 3':'Cost',
                                'Unnamed: 4':'Market Value',
                                'Unnamed: 5':'Unreal PNL',
                                'Unnamed: 6':'Requirement',
                                'Unnamed: 7':'Real PNL'}, inplace=True)
Daily_Change_x.reset_index(inplace = True)
Daily_Change_y.rename(columns={'Account Number': 'Inventory Sub Totals',
                                'Total Available Funds':'Account Name',
                                'Unnamed: 2':'Position Type',
                                'Unnamed: 3':'Cost',
                                'Unnamed: 4':'Market Value',
                                'Unnamed: 5':'Unreal PNL',
                                'Unnamed: 6':'Requirement',
                                'Unnamed: 7':'Real PNL'}, inplace=True)

Daily_Change_y.reset_index(inplace = True)
Daily_Change_x=pd.merge(Daily_Change_x, Daily_Change_y, on='index', how='left')
Daily_Change_x['Cost'] = Daily_Change_x['Cost_x']-Daily_Change_x['Cost_y']
Daily_Change_x['Market Value']=Daily_Change_x['Market Value_x']-Daily_Change_x['Market Value_y']
Daily_Change_x['Requirement']=Daily_Change_x['Requirement_x']-Daily_Change_x['Requirement_y']
Daily_Change_x['Unreal PNL']=Daily_Change_x['Unreal PNL_x']-Daily_Change_x['Unreal PNL_y']
Daily_Change_x['Real PNL']=Daily_Change_x['Real PNL_x']-Daily_Change_x['Real PNL_y']
Daily_Change_x = Daily_Change_x[['index','Account Name_x','Position Type_x','Cost','Market Value','Requirement','Unreal PNL','Real PNL']]
Daily_Change = Daily_Change_x.groupby('Account Name_x')['Cost','Market Value','Requirement', 'Unreal PNL','Real PNL'].sum()
Daily_Change['Account Name_x']=['K72 Muni Inv','K74 Corporates ','K76 S P Inv',
                                'K77 CD Inv','K78 Taxable Mun',
                                 'K79 Cali Tax Ex','K80 Muni Tax Ex',
                                'K81 Muni Tax','K82 Tax 0 Muni','L81 Sierra Comp',
                                  'M64 Sierra MBS','N90 CD','P01 Corp Floate','P02 Corp Sp','N88 Corp Notes']
Daily_Change = Daily_Change_x.append(Daily_Change, ignore_index = True,sort = False)
Daily_Change['Account Name_x'] = Daily_Change['Account Name_x'].replace({'K72 Muni Inv':'K72 Muni Inv Fl'})
Daily_Change.sort_values(['Account Name_x','Position Type_x'],inplace = True)
Daily_Change.rename(columns={'Account Name_x':'Account Name','Position Type_x':'Position Type'},inplace = True)
Daily_Change.drop('index',axis = 1, inplace = True)

"""
Create and Format Summary sheet

# """
Hilltop_Summary_new = pd.read_excel(io=Recent, sheet_name='Detail')
Hilltop_Summary_new = Hilltop_Summary_new.groupby(['Account Name']).agg({
                                                          'Unreal PNL':'sum',
                                                          'Real PNL':'sum',
                                                          'Requirement':'sum',
                                                          'Cost':'sum',
                                                          'Market Value':'sum'})
Hilltop_Summary_new = Hilltop_Summary_new[['Cost','Market Value','Requirement','Unreal PNL','Real PNL']]

Hilltop_Summary_new.loc['Column_Total'] = Hilltop_Summary_new.sum(numeric_only=True, axis=0)
Daily_Change.drop([1,5,7,9,11,13,15,17,19,21,33,35,36,37,38,39,40,41,42,43],inplace = True)
Daily_Change = Daily_Change.reindex([0,4,6,10,12,14,16,18,20,22,23,47,24,25,44,26,27,45,28,29,46,2,3,34,30,31,32])
Daily_Change.to_excel(writer,sheet_name ='Summary',index=False,startrow=20)
Hilltop_Summary_new.to_excel(writer,sheet_name ='Summary',index=True,startrow=1)

Hilltop_Recent_s['Change'] = Hilltop_Recent_s['Total Available Funds']-Hilltop_Old_s['Total Available Funds']
Hilltop_Recent_s = Hilltop_Recent_s[3:]
Hilltop_Recent_s = Hilltop_Recent_s[['Account Number','Total Available Funds','Change']]
Hilltop_Recent_s.to_excel(writer,sheet_name ='Summary',index=False,startrow=1,startcol=8)
workbook = writer.book

worksheet_summary = writer.sheets['Summary']
format1 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5'})
format2 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9',
                               'bg_color':'#f0f0f0'})
format3 = workbook.add_format()
format4 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5',
                               'bold': True,
                               'bg_color':'#f0f0f0'})
format5 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5',
                               'align':'center',
                               'bold': True,'bg_color':'#5551b8',
                               'font_color':'white'})
format7 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9',
                               'bg_color':'#f0f0f0'})
format7 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9'})
format8 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'11',
                               'bold': True})
format111 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'10',
                                'bg_color':'#f0f0f0',
                                'bold': True})
format_1 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'10',
                                'bg_color':'#f0f0f0'})
format_11 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'10',
                                'bold': True})
worksheet_summary.write('A2', 'Account Name',format111)
worksheet_summary.write('B2', 'Cost',format111)
worksheet_summary.write('C2', 'Market Value',format111)
worksheet_summary.write('D2', 'Requirment',format111)
worksheet_summary.write('E2','Unreal PNL',format111)
worksheet_summary.write('F2','Real PNL',format111)

worksheet_summary.write('I21', 'Account Name',format111)
worksheet_summary.write('J21', 'Cost',format111)
worksheet_summary.write('K21', 'Market Value',format111)
worksheet_summary.write('L21', 'Requirment',format111)
worksheet_summary.write('M21','Unreal PNL',format111)
worksheet_summary.write('N21','Real PNL',format111)

worksheet_summary.write('I2', 'Account Name',format111)
worksheet_summary.write('J2', 'Total Funds',format111)
worksheet_summary.write('K2', 'Change',format111)

worksheet_summary.write('A21', 'Account Name',format111)
worksheet_summary.write('B21', 'Position Type',format111)
worksheet_summary.write('C21', 'Cost',format111)
worksheet_summary.write('D21', 'Market Value',format111)
worksheet_summary.write('E21', 'Requirment',format111)
worksheet_summary.write('F21', 'Unreal PNL',format111)
worksheet_summary.write('G21', 'Real PNL',format111)

worksheet_summary.write('A18', 'Total',format5)
worksheet_summary.write('B46', 'Total Long',format5)
worksheet_summary.write('B47', 'Total Short',format5)
worksheet_summary.write('B48', 'Total',format5)

worksheet_summary.write('A3', 'K72 Muni Inv Fl',format_1)
worksheet_summary.write('A4', 'K74 Corporates',format_1)
worksheet_summary.write('A5', 'K76 S P Inv',format_1)
worksheet_summary.write('A6', 'K77 CD Inv',format_1)
worksheet_summary.write('A7','K78 Taxable Mun',format_1)
worksheet_summary.write('A8', 'K79 Cali Tax Ex',format_1)
worksheet_summary.write('A9', 'K80 Muni Tax Ex',format_1)
worksheet_summary.write('A10', 'K81 Muni Tax',format_1)
worksheet_summary.write('A11', 'K82 Tax 0 Muni',format_1)
worksheet_summary.write('A12','L81 Sierra Comp',format_1)
worksheet_summary.write('A13', 'M64 Sierra MBS',format_1)
worksheet_summary.write('A14', 'N90',format_1)
worksheet_summary.write('A15', 'P01 Corp Floate',format_1)
worksheet_summary.write('A16', 'P02 Corp SP',format_1)
worksheet_summary.write('A17', 'N88 Corp Notes',format_1)

worksheet_summary.write('A22', 'K72 Muni Inv Fl',format_1)
worksheet_summary.write('A23', 'K76 S P Inv',format_1)
worksheet_summary.write('A24', 'K77 CD Inv',format_1)
worksheet_summary.write('A25', 'K79 Cali Tax Ex',format_1)
worksheet_summary.write('A26', 'K80 Muni Tax Ex',format_1)
worksheet_summary.write('A27', 'K81 Muni Tax',format_1)
worksheet_summary.write('A28', 'K82 Tax 0 Muni',format_1)
worksheet_summary.write('A29', 'L81 Sierra Comp',format_1)
worksheet_summary.write('A30', 'M64 Sierra MBS',format_1)
worksheet_summary.write('A31', 'N88 Corp Notes',format_1)
worksheet_summary.write('A32', ' ',format_1)
worksheet_summary.write('A33', ' ',format_1)
worksheet_summary.write('A34', 'N90 CD ',format_1)
worksheet_summary.write('A35', ' ',format_1)
worksheet_summary.write('A36', ' ',format_1)
worksheet_summary.write('A37', 'P01 Corp Floate',format_1)
worksheet_summary.write('A38', ' ',format_1)
worksheet_summary.write('A39', ' ',format_1)
worksheet_summary.write('A40', 'P02 Corp SP',format_1)
worksheet_summary.write('A41', ' ',format_1)
worksheet_summary.write('A42', ' ',format_1)
worksheet_summary.write('A43', 'K74 Corp',format_1)
worksheet_summary.write('A44', ' ',format_1)
worksheet_summary.write('A45', ' ',format_1)

worksheet_summary.write('B22', 'Total',format_11)
worksheet_summary.write('B23', 'Total',format_11)
worksheet_summary.write('B24', 'Total',format_11)
worksheet_summary.write('B25', 'Total',format_11)
worksheet_summary.write('B26', 'Total',format_11)
worksheet_summary.write('B27', 'Total',format_11)
worksheet_summary.write('B28', 'Total',format_11)
worksheet_summary.write('B29', 'Total',format_11)
worksheet_summary.write('B30', 'Total',format_11)
worksheet_summary.write('B33', 'Total',format_11)
worksheet_summary.write('B36', 'Total',format_11)
worksheet_summary.write('B39', 'Total',format_11)
worksheet_summary.write('B42', 'Total',format_11)
worksheet_summary.write('B45', 'Total',format_11)


worksheet_summary.write('I3', 'Cost',format_1)
worksheet_summary.write('I4', 'Market Value',format_1)
worksheet_summary.write('I5', 'Unreal PNL',format_1)
worksheet_summary.write('I6', 'Real PNL',format_1)
worksheet_summary.write('I7', 'Requirement',format_1)
worksheet_summary.write('I8', 'Available Funds',format_1)
worksheet_summary.write('I9', 'Call/Excess',format_1)

worksheet_summary.set_column('A1:A14', 15) #2
worksheet_summary.set_column('B1:G14', 12, format1) #2

worksheet_summary.set_column('H1:H19',8,format1)
worksheet_summary.set_column('I1:I15',15,format1)
worksheet_summary.set_column('J1:N15',12,format1)

worksheet_summary.write('I22', 'Muni',format_1)
worksheet_summary.write_formula('J22','=SUM(C22:C23,C25:C29)')
worksheet_summary.write_formula('K22','=SUM(C22:C23,C25:C29)')
worksheet_summary.write_formula('L22','=SUM(C22:C23,C25:C29)')
worksheet_summary.write_formula('M22','=SUM(C22:C23,C25:C29)')
worksheet_summary.write_formula('N22','=SUM(C22:C23,C25:C29)')

worksheet_summary.write('I23', 'Corp',format_1)
worksheet_summary.write_formula('J23','=SUM(C33,C39,C42,C45)')
worksheet_summary.write_formula('K23','=SUM(C33,C39,C42,C45)')
worksheet_summary.write_formula('L23','=SUM(C33,C39,C42,C45)')
worksheet_summary.write_formula('M23','=SUM(C33,C39,C42,C45)')
worksheet_summary.write_formula('N23','=SUM(C33,C39,C42,C45)')

worksheet_summary.write('I24', 'CD',format_1)
worksheet_summary.write_formula('J24','=SUM(C36)')
worksheet_summary.write_formula('K24','=SUM(D36)')
worksheet_summary.write_formula('L24','=SUM(E36)')
worksheet_summary.write_formula('M24','=SUM(F36)')
worksheet_summary.write_formula('N24','=SUM(G36)')

worksheet_summary.write('I25', 'CMO',format_1)
worksheet_summary.write_formula('J25','=SUM(C30)')
worksheet_summary.write_formula('K25','=SUM(D30)')
worksheet_summary.write_formula('L25','=SUM(E30)')
worksheet_summary.write_formula('M25','=SUM(F30)')
worksheet_summary.write_formula('N25','=SUM(G30)')

worksheet_summary.set_zoom(95)

# worksheet_summary.set_column('L:P', None, None, {'hidden': True})

merge_format = workbook.add_format({
    'bold': 1,
    'border': 0,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#5551b8',
    'font_color':'white'})


worksheet_summary.merge_range('A1:F1', 'Monthly Summary', merge_format)
worksheet_summary.merge_range('A20:G20', 'Daily Change', merge_format)
worksheet_summary.merge_range('I20:N20', 'Product Summary',merge_format)
worksheet_summary.merge_range('I1:K1', 'Summary',merge_format)

"""
Create and format Detail sheet
"""

Hilltop_Recent.to_excel(writer, sheet_name='Detail', index=False)


worksheet = writer.sheets['Detail']

format1 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5',
                               'bottom':True})
format2 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9',
                               'bg_color':'#f0f0f0',
                               'bottom':True})
format3 = workbook.add_format()
format4 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5',
                               'bold': True,
                               'bottom':True,
                               'bg_color':'#f0f0f0'})
format5 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5',
                               'align':'center',
                               'bold': True,'bg_color':'#5551b8',
                               'font_color':'white',
                               'bottom':True})
format7 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9',
                               'bg_color':'#f0f0f0',
                               'bottom':True})
format6=workbook.add_format({'bottom':False})

format8=workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5',
                               'font_color':'green'})
format9=workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5',
                               'font_color':'red'})
worksheet_summary.conditional_format('K5:K6', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('K5:K6', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format9})
worksheet_summary.conditional_format('K8:K9', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('K8:K9', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format9})
worksheet_summary.conditional_format('F22:G65', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('F22:G65', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format9})

worksheet_summary.conditional_format('E3:F18', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('E3:F18', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format9})
worksheet_summary.conditional_format('M22:N27', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('M22:N27', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format9})
"""Requirement Change """
worksheet_summary.conditional_format('K7', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('K7', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format9})
worksheet_summary.conditional_format('E22:E65', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('E22:E65', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format9})

worksheet_summary.conditional_format('D3:D18', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('D3:D18', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format9})
worksheet_summary.conditional_format('L22:L27', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('L22:L27', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format9})

worksheet_summary.insert_image('M1', 'P:/1. Individual Folders/Chad/Python Scripts/PNL Report/Logo.png',{'x_scale':0.5,'y_scale':0.5})

format1.set_text_wrap(True)
format3.set_bg_color('#cfcfcf')
format4.set_text_wrap(True)
format2.set_text_wrap(True)
format5.set_text_wrap(True)
format1.set_bottom(4)
format3.set_bottom(4)
format4.set_bottom(4)
format2.set_bottom(4)
format7.set_bottom(4)
# Format each colum to fit and display data correclty

worksheet.set_column('A:A', 24, format7) #2
worksheet.set_column('B:B', 9.5, format2) #2
worksheet.set_column('C:C', 14, format2) #3
worksheet.set_column('D:D', 11, format1)#4
worksheet.set_column('E:E', 11, format1)#5
worksheet.set_column('F:F', 11, format1)#6
worksheet.set_column('G:G', 11, format1)#7
worksheet.set_column('H:H', 11, format1)#8
worksheet.set_column('I:I', 11, format1)#9
worksheet.set_column('J:J', 11, format1)#10
worksheet.set_column('K:K', 11, format1)#11
worksheet.set_column('L:L', 11, format1)#12
worksheet.set_column('M:M', 11, format1)#13
worksheet.set_column('N:N', 11, format1)#14
worksheet.set_column('O:O', 11, format1)#15

worksheet.write('A1', 'Security',format5)
worksheet.write('B1', 'Cusip',format5)
worksheet.write('C1', 'Account',format5)
worksheet.write('D1', 'Price',format5)
worksheet.write('E1', 'TW QTY ',format5)
worksheet.write('F1', 'HT QTY',format5)
worksheet.write('G1', 'QTY Discrepancy',format5)
worksheet.write('H1', 'HT QTY Change',format5)
worksheet.write('I1', 'HT New Unreal PNL',format5)
worksheet.write('J1', 'HT Old Unreal PNL',format5)
worksheet.write('K1', 'Real PNL Change',format5)
worksheet.write('L1', 'Adj Unreal PNL Change',format5)
worksheet.write('M1', 'TW PNL',format5)
worksheet.write('N1', 'HT-TW PNL Discrep.',format5)
worksheet.write('O1', 'Requirement Change',format5)

worksheet.set_zoom(90)


QTY_Discrepancy = 'Discrepancy Between TW Quantity & Hilltop Quantity.'
HT_QTY_Change = 'Amount bought or sold on the day.'
Adj_Unreal_PNL_Change = 'Difference between Hilltop Unreal PNL and Real PNL. \n\nShows overall profit or loss while adjusting for the change in unrealized PNL. \n\nFor example: Bond has 1000 in Unreal PNL, it sells for a gain of 750. There would be a -250 change'
HT_TW_PNL_Discrepancy = 'Discrepancy between Hilltop  Real PNL and TW Real PNL.'
Requirement_Change = 'Increase or decrease in requirement.'

worksheet.write_comment('G1', QTY_Discrepancy, {'visible': 0})
worksheet.write_comment('H1', HT_QTY_Change, {'visible': 0})
worksheet.write_comment('L1', Adj_Unreal_PNL_Change, {'visible': 0,'x_scale': 1.5, 'y_scale': 1.8})
worksheet.write_comment('N1', HT_TW_PNL_Discrepancy, {'visible': 0})
worksheet.write_comment('O1', Requirement_Change, {'visible': 0})

worksheet.freeze_panes(1, 1)
worksheet.autofilter('A1:O20000')
worksheet.hide_gridlines(2)




Hilltop_x = Hilltop_Individual_Summary
Hilltop_y = Hilltop_x
# for cusip in Hilltop_Individual_Summary:
#     if Hilltop_Individual_Summary.loc['Security'] == 0:
#         Hilltop_Individual_Summary.loc['Cusip']['Security'] = Hilltop_Recent_x.loc['Cusip']['Description']

def Summary_Individual_Sheets(Hilltop_Individual_Summary):
    column_summary = ('TW - HT Quantity Discrepancy',
                      'HT Change in Quantity',
                      'Adj Unreal PNL Change',
                      'HT-TW PNL Discrepancy',
                      'Requirement Change')
    for item in column_summary:
        Hilltop_Individual_Summary[item] = Hilltop_Individual_Summary[item]
    Hilltop_QTY_DSP = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['TW - HT Quantity Discrepancy'] != 0)]
    Hilltop_HT_QTY_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['HT Change in Quantity'] != 0)]
    Hilltop_Adj_Unreal_PNL_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['Adj Unreal PNL Change'] != 0)]
    Hilltop_HT_TW_PNL_DSP = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['HT-TW PNL Discrepancy'] != 0)]
    Hilltop_Requirement_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['Requirement Change'] != 0)]
    return Hilltop_QTY_DSP,Hilltop_HT_QTY_Change,Hilltop_Adj_Unreal_PNL_Change, Hilltop_HT_TW_PNL_DSP,Hilltop_Requirement_Change

Individual_Sheets = Summary_Individual_Sheets(Hilltop_Individual_Summary)


Hilltop_QTY_DSP  = Individual_Sheets[0]
Hilltop_QTY_DSP = pd.merge(Hilltop_QTY_DSP, Hilltop_x, on='Cusip', how='left')
Hilltop_QTY_DSP['TW - HT Quantity Discrepancy_x'] = Hilltop_QTY_DSP['TW - HT Quantity Discrepancy_x'].abs()
Hilltop_QTY_DSP.sort_values('TW - HT Quantity Discrepancy_x', axis=0, ascending=False, inplace=True)
Hilltop_QTY_DSP = Hilltop_QTY_DSP[['Security_x','Cusip','Account_x','TW - HT Quantity Discrepancy_y']]
Hilltop_QTY_DSP.rename(columns={'Security_x': 'Security','Account_x':'Account','TW - HT Quantity Discrepancy_y':"QTY DSP"}, inplace=True)
QTY_DSP_Cleared_Positions_Drop = QTY_DSP_Cleared_Positions.drop('Position Notes',axis = 1)
Hilltop_QTY_DSP = Hilltop_QTY_DSP.append(QTY_DSP_Cleared_Positions_Drop)
Hilltop_QTY_DSP.drop_duplicates(subset ='Cusip',keep = False, inplace = True)
Hilltop_QTY_DSP = Hilltop_QTY_DSP[['Security','Account','Cusip','QTY DSP']]
Hilltop_QTY_DSP.to_excel(writer, sheet_name = 'HT QTY DSP', index=False)
worksheet_Hilltop_QTY_DSP = writer.sheets['HT QTY DSP']
QTY_DSP_Cleared_Positions= QTY_DSP_Cleared_Positions[['Security','Account','Cusip','QTY DSP','Position Notes']]
QTY_DSP_Cleared_Positions.to_excel(writer,sheet_name ='HT QTY DSP',index=False,startrow=1,startcol=6)

Hilltop_HT_TW_PNL_DSP = Individual_Sheets[3]
Hilltop_HT_TW_PNL_DSP.sort_values('HT-TW PNL Discrepancy', axis=0, ascending=False, inplace=True)
Hilltop_HT_TW_PNL_DSP = pd.merge(Hilltop_HT_TW_PNL_DSP, Hilltop_x, on='Cusip', how='left')
Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_x'] = Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_x'].abs()
Hilltop_HT_TW_PNL_DSP.sort_values('HT-TW PNL Discrepancy_x', axis=0, ascending=False, inplace=True)
Hilltop_HT_TW_PNL_DSP = Hilltop_HT_TW_PNL_DSP[['Security_x','Cusip','Account_x','HT-TW PNL Discrepancy_y']]
Hilltop_HT_TW_PNL_DSP = Hilltop_HT_TW_PNL_DSP[(Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_y'] > 10)]
Hilltop_HT_TW_PNL_DSP.to_excel(writer, sheet_name = 'HT-TW PNL DSP', index=False)
worksheet_Hilltop_HT_TW_PNL_DSP = writer.sheets['HT-TW PNL DSP']


Hilltop_Adj_Unreal_PNL_Change = Individual_Sheets[2]
Hilltop_Adj_Unreal_PNL_Change = pd.merge(Hilltop_Adj_Unreal_PNL_Change, Hilltop_x, on='Cusip', how='left')
Hilltop_Adj_Unreal_PNL_Change['Adj Unreal PNL Change_x'] = Hilltop_Adj_Unreal_PNL_Change['Adj Unreal PNL Change_x'].abs()
Hilltop_Adj_Unreal_PNL_Change.sort_values('Adj Unreal PNL Change_x', axis=0, ascending=False, inplace=True)
Hilltop_Adj_Unreal_PNL_Change = Hilltop_Adj_Unreal_PNL_Change[['Security_x','Cusip','Account_x','Adj Unreal PNL Change_y']]
Hilltop_Adj_Unreal_PNL_Change.to_excel(writer, sheet_name = 'Adj Unreal PNL Change', index=False)
worksheet_Hilltop_Adj_Unreal_PNL_Change = writer.sheets['Adj Unreal PNL Change']


Hilltop_Requirement_Change = Individual_Sheets[4]
Hilltop_Requirement_Change = pd.merge(Hilltop_Requirement_Change, Hilltop_x, on='Cusip', how='left')
Hilltop_Requirement_Change = Hilltop_Requirement_Change[['Cusip','Security_x','Account_x','Requirement Change_x','Requirement Change_y']]
Hilltop_Requirement_Change['Requirement Change_x'] = Hilltop_Requirement_Change['Requirement Change_x'].abs()
Hilltop_Requirement_Change.sort_values('Requirement Change_x', axis=0, ascending=False, inplace=True)
Hilltop_Requirement_Change = Hilltop_Requirement_Change[['Security_x','Cusip','Account_x','Requirement Change_y']]
Hilltop_Requirement_Change.to_excel(writer, sheet_name = 'Requirement Change', index=False)
worksheet_Hilltop_Requirement_Change = writer.sheets['Requirement Change']

worksheet_Hilltop_QTY_DSP.set_column('A:A', 20, format7) #2
worksheet_Hilltop_QTY_DSP.set_column('B:B', 12, format2) #2
worksheet_Hilltop_QTY_DSP.set_column('C:C', 12, format7) #3
worksheet_Hilltop_QTY_DSP.set_column('D:D', 12, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('D:D', 12, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('E:E', 30, format1)

worksheet_Hilltop_QTY_DSP.write('A1', 'Security',format5)
worksheet_Hilltop_QTY_DSP.write('B1', 'Account',format5)
worksheet_Hilltop_QTY_DSP.write('C1', 'Cusip',format5)
worksheet_Hilltop_QTY_DSP.write('D1', 'QTY DSP',format5)
worksheet_Hilltop_QTY_DSP.write('E1', 'Position Notes',format5)

worksheet_Hilltop_QTY_DSP.write('G2', 'Security',format5)
worksheet_Hilltop_QTY_DSP.write('H2', 'Account',format5)
worksheet_Hilltop_QTY_DSP.write('I2', 'Cusip',format5)
worksheet_Hilltop_QTY_DSP.write('J2', 'QTY DSP',format5)
worksheet_Hilltop_QTY_DSP.write('K2', 'Position Notes',format5)

worksheet_Hilltop_QTY_DSP.merge_range('G1:K1', 'Cleared QTY DSP',merge_format)

worksheet_Hilltop_QTY_DSP.set_column('G:G', 20, format1) #2
worksheet_Hilltop_QTY_DSP.set_column('H:H', 12, format1) #2
worksheet_Hilltop_QTY_DSP.set_column('I:I', 12, format1) #3
worksheet_Hilltop_QTY_DSP.set_column('J:J', 15, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('K:K', 30, format1)#4

worksheet_Hilltop_QTY_DSP.freeze_panes(1, 1)
worksheet_Hilltop_QTY_DSP.autofilter('A1:V20000')
worksheet_Hilltop_QTY_DSP.hide_gridlines(2)
worksheet_Hilltop_QTY_DSP.protect('welcome123')
worksheet_Hilltop_QTY_DSP.set_zoom(90)
"""
"""
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('A:A', 30, format7) #2
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('B:B', 12, format2) #2
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('C:C', 12, format7) #3
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('D:D', 12, format1)#4

worksheet_Hilltop_Adj_Unreal_PNL_Change.write('A1', 'Security',format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('B1', 'Cusip',format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('C1', 'Account',format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('D1', 'Adj Unreal PNL Change ',format5)
"""
"""
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('A:A', 30, format7) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('B:B',  9, format2) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('C:C', 12, format7) #3
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('D:D', 11, format1)#4

worksheet_Hilltop_HT_TW_PNL_DSP.write('A1', 'Security',format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('B1', 'Cusip',format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('C1', 'Account',format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('D1', 'PNL DSP ',format5)

"""
"""
worksheet_Hilltop_Requirement_Change.set_column('A:A', 30, format7) #2
worksheet_Hilltop_Requirement_Change.set_column('B:B', 12, format2) #2
worksheet_Hilltop_Requirement_Change.set_column('C:C', 12, format7) #3
worksheet_Hilltop_Requirement_Change.set_column('D:D', 12, format1)#4

worksheet_Hilltop_Requirement_Change.write('A1', 'Security',format5)
worksheet_Hilltop_Requirement_Change.write('B1', 'Cusip',format5)
worksheet_Hilltop_Requirement_Change.write('C1', 'Account',format5)
worksheet_Hilltop_Requirement_Change.write('D1', 'Requirement Change',format5)



worksheet_Hilltop_QTY_DSP.freeze_panes(1, 1)
worksheet_Hilltop_QTY_DSP.autofilter('A1:E20000')
worksheet_Hilltop_QTY_DSP.hide_gridlines(2)

worksheet_Hilltop_HT_TW_PNL_DSP.freeze_panes(1, 1)
worksheet_Hilltop_HT_TW_PNL_DSP.autofilter('A1:D20000')
worksheet_Hilltop_HT_TW_PNL_DSP.hide_gridlines(2)


worksheet_Hilltop_Adj_Unreal_PNL_Change.freeze_panes(1, 1)
worksheet_Hilltop_Adj_Unreal_PNL_Change.autofilter('A1:D20000')
worksheet_Hilltop_Adj_Unreal_PNL_Change.hide_gridlines(2)

worksheet_Hilltop_Requirement_Change.freeze_panes(1, 1)
worksheet_Hilltop_Requirement_Change.autofilter('A1:D20000')
worksheet_Hilltop_Requirement_Change.hide_gridlines(2)

worksheet_summary.hide_gridlines(2)

workbook.close()

# import win32com.client
# from win32com.client import Dispatch, constants
# const=win32com.client.constants

# olMailItem = 0x0
# obj = win32com.client.Dispatch("Outlook.Application")
# newMail = obj.CreateItem(olMailItem)
# newMail.Subject = "PNL Report" 
# newMail.To = 'ccraig90@bloomberg.net'#; lankowsky@bloomberg.net; jdean18@bloomberg.net; jblamire@bloomberg.net'
# newMail.HTMLBody = 'A new PNL Report has been created.'
# # newMail.Attachments.Add('P:/2. Corps/PNL_Daily_Report/Reports/'+ str(month) + str('.') + str(day) + str('.') + str(year) + str('PNL Discrepancy.xlsx'))
# newMail.display()
# newMail.Send()









def subject_name(file_text):
    now = datetime.datetime.now()
    weekday = now.weekday()
    if weekday == 0:
        today = now - datetime.timedelta(days=3)
        yesterday = today - datetime.timedelta(days=1)
    elif weekday == 1:
        today = now - datetime.timedelta(days=1)
        yesterday = today - datetime.timedelta(days=3)
    else:
        today = now - datetime.timedelta(days=1)
        yesterday = today - datetime.timedelta(days=1)
    current = today.strftime("%Y-%m-%d")
    current = str(file_text)+str(current)
    PNL_Report_Date = today.strftime("%m.%d.%Y")
    PNL_Report_Date = str(PNL_Report_Date)+'PNL Discrepancy'
    now = datetime.datetime.now()
    PNL_Report_Write_to_Date = now.strftime('%m.%d.%Y')
    PNL_Report_Write_to_Date = str(PNL_Report_Write_to_Date)+'PNL Discrepancy'
    today = file_text + str(today)
    yesterday = yesterday.strftime("%Y-%m-%d")
    yesterday = file_text + str(yesterday)
    return current,yesterday,PNL_Report_Date,PNL_Report_Write_to_Date

def correct_position_type(Inventory):
    x = Inventory['Position']
    x = pd.to_numeric(x)
    Inventory['Position'] = x
    x = Inventory['P&L']
    x = pd.to_numeric(x)
    Inventory['P&L'] = x
    return Inventory


file_text = 'Inventory Margin Report for '
x = subject_name(file_text)

"""
Pull and Generate new QTY_DSP_Cleared_Positions file
"""
PNL_Report_Date = x[2]
PNL_Report_File_Most_Recent = 'P:/2. Corps/PNL_Daily_Report/Reports/PNL_Report.xlsx'
Cleared_Yesterday_PNL_Report = pd.read_excel(PNL_Report_File_Most_Recent,sheet_name = 'HT QTY DSP')
Cleared_Yesterday_PNL_Report = Cleared_Yesterday_PNL_Report[['Security','Account','Cusip','QTY DSP','Position Notes']]
Cleared_Yesterday_PNL_Report.dropna(inplace = True)
QTY_DSP_Cleared_Positions = 'P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/QTY_DSP_Cleared_Positions.xlsx'
QTY_DSP_Cleared_Positions= pd.read_excel(QTY_DSP_Cleared_Positions,index = False)
QTY_DSP_Cleared_Positions = QTY_DSP_Cleared_Positions.append(Cleared_Yesterday_PNL_Report)

QTY_DSP_Cleared_Positions.drop_duplicates(keep='first',inplace = True)
writer = pd.ExcelWriter('P:/2. Corps/PNL_Daily_Report/Cleared_Position_File/QTY_DSP_Cleared_Positions.xlsx', engine='xlsxwriter')
QTY_DSP_Cleared_Positions.to_excel(writer)
writer.save()


today = x[0]
yesterday = x[1]
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()

i = 0
while i < 20:
    if message.Subject == today:
        try:
            attachments = message.Attachments
            attachment = attachments.Item(1)
            attachment.SaveASFile('P:/2. Corps/PNL_Daily_Report/HT_Files/' + str(today)+'.xlsx') #Saves to the attachment to current folder
            print('HT File Found')
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()


file_text = 'Report "TW 16 22" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_16_22 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22.transpose()
Bloomberg_Inventory_16_22.reset_index(inplace = True)
Bloomberg_Inventory_16_22.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_16_22.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22.transpose()
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[:-1]
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[2:]
Bloomberg_Inventory_16_22 = Bloomberg_Inventory_16_22[0].str.split(',',expand=True)
Bloomberg_Inventory_16_22.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_16_22.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_16_22.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_16_22 = correct_position_type(Bloomberg_Inventory_16_22)

file_text = 'Report "TW 1 5" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:

    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_1_5 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5.transpose()
Bloomberg_Inventory_1_5.reset_index(inplace = True)
Bloomberg_Inventory_1_5.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_1_5.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5.transpose()
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[:-1]
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[2:]
Bloomberg_Inventory_1_5 = Bloomberg_Inventory_1_5[0].str.split(',',expand=True)
Bloomberg_Inventory_1_5.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_1_5.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_1_5.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_1_5 = correct_position_type(Bloomberg_Inventory_1_5)

file_text = 'Report "TW 6 10" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
    
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_6_10 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10.transpose()
Bloomberg_Inventory_6_10.reset_index(inplace = True)
Bloomberg_Inventory_6_10.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_6_10.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10.transpose()
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[:-1]
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[2:]
Bloomberg_Inventory_6_10 = Bloomberg_Inventory_6_10[0].str.split(',',expand=True)
Bloomberg_Inventory_6_10.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_6_10.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_6_10.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_6_10 = correct_position_type(Bloomberg_Inventory_6_10)

file_text = 'Report "TW 11 15" is generated.'
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the inbox. You can change that number to reference
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
i = 0
while i < 10:
   
    if message.Subject == file_text:
        print(message.Subject)
        try:
            attachments = message.Attachments
            attachment = attachments.Item(2)
            attachment.SaveASFile('C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\Bloomberg TW'+'.xls') #Saves to the attachment to current folder
            message = messages.GetNext()

        except:
            message = messages.GetNext()
    i += 1
    message = messages.GetNext()
    
Bloomberg_Inventory_11_15 = pd.read_table('C:/Users/ccraig/Desktop/PNL Project/Bloomberg TW.xls')
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15.transpose()
Bloomberg_Inventory_11_15.reset_index(inplace = True)
Bloomberg_Inventory_11_15.drop('index',axis = 1, inplace = True)
Bloomberg_Inventory_11_15.drop(0,axis = 1, inplace = True)
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15.transpose()
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[:-1]
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[2:]
Bloomberg_Inventory_11_15 = Bloomberg_Inventory_11_15[0].str.split(',',expand=True)
Bloomberg_Inventory_11_15.rename(columns={0: 'Security'}, inplace=True)
Bloomberg_Inventory_11_15.rename(columns={1: 'P&L'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={2: 'Cusip'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={3: 'Position'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={4: 'Symbol'},inplace =True)
Bloomberg_Inventory_11_15.rename(columns={5: 'Book'},inplace =True)
Bloomberg_Inventory_11_15 = correct_position_type(Bloomberg_Inventory_11_15)

Bloomberg_Inventory = pd.concat([Bloomberg_Inventory_16_22,
                                 Bloomberg_Inventory_1_5,
                                 Bloomberg_Inventory_6_10,
                                 Bloomberg_Inventory_11_15],ignore_index = True)

"""
Read in Excel files from Hilltop and Bloomberg

"""
# Bloomberg_Inventory = pd.read_excel('C:/Users/ccraig/Desktop/PNL Project/8.21.2019 Inventory TW.xlsx')
# x = Bloomberg_Inventory[4:]
# x.rename(columns={x.columns[0]: 'Security',x.columns[1]:'P&L',x.columns[2]:'Cusip',x.columns[3]:'Position',x.columns[4]:'Symbol'}, inplace=True)
# Bloomberg_Inventory.rename(columns={'CUSIP': 'Cusip'}, inplace=True)
Bloomberg_Inventory['Cusip'] = Bloomberg_Inventory['Symbol'].str[1:10]
Bloomberg_Inventory = Bloomberg_Inventory[['Cusip', 'P&L', 'Security', 'Position','Symbol','Book']]
Bloomberg_Inventory['Position'] = Bloomberg_Inventory['Position']*1000
# Bloomberg_Inventory['Cusip'] = Bloomberg_Inventory['Cusip'].astype(str)
print(today)
print(yesterday)
Recent = 'C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\' + str(today)+'.xlsx'#Recent = 'C:/Users/ccraig/Desktop/PNL Project/'+str(today)+'.xlsx'
Old = 'C:\\Users\\ccraig\\Desktop\\PNL Project' + '\\' + str(yesterday)+'.xlsx'
Hilltop_Recent_x = pd.read_excel(io=Recent, sheet_name='Detail')
# Hilltop_Recent_x['Cusip'] = Hilltop_Recent['Cusip'].astype(str)
Hilltop_Old_y = pd.read_excel(io=Old, sheet_name='Detail')
# Hilltop_Old_y['Cusip'] = Hilltop_Old_y['Cusip'].astype(str)
Hilltop_Recent_s = pd.read_excel(io=Recent, sheet_name='Summary')
Hilltop_Old_s = pd.read_excel(io=Old, sheet_name='Summary')
Hilltop_Recent_s = Hilltop_Recent_s.head(10)
Hilltop_Old_s = Hilltop_Old_s.head(10)
Hilltop_Recent_x['Cusip_group_by'] = Hilltop_Recent_x['Cusip']
Hilltop_Recent_x['Cusip_group_by'] = 'C'+ Hilltop_Recent_x['Cusip_group_by']
Bloomberg_Inventory = Bloomberg_Inventory.groupby(['Cusip']).agg({'P&L':'sum',
                                                                   'Security':'first',
                                                                   'Position':'sum',
                                                                   'Symbol':'first',
                                                                 'Book':'first'})


Hilltop_Recent = Hilltop_Recent_x.groupby(['Cusip_group_by']).agg({'Quantity':'sum',
                                                                   'Unreal PNL':'sum',
                                                                   'Real PNL':'sum',
                                                                   'Requirement':'sum',
                                                                   'Cusip':'first',
                                                                   'Description':'first',
                                                                   'Price':'mean'})

Hilltop_Old_y['Cusip_group_by'] = Hilltop_Old_y['Cusip']
Hilltop_Old = Hilltop_Old_y.groupby(['Cusip_group_by']).agg({'Quantity':'sum',
                                                             'Unreal PNL':'sum',
                                                             'Real PNL':'sum',
                                                             'Requirement':'sum',
                                                             'Cusip':'first',
                                                             'Description':'first',
                                                             'Price':'mean'})

"""
Merge Hilltop Recent and Old together

"""

Hilltop_Recent = pd.merge(Hilltop_Recent, Hilltop_Old, on='Cusip', how='left')

Hilltop_Recent = pd.merge(Hilltop_Recent, Bloomberg_Inventory, on='Cusip', how='outer')

Hilltop_Recent_x = Hilltop_Recent_x[['Cusip','Account Name']]

Hilltop_Recent = pd.merge(Hilltop_Recent, Hilltop_Recent_x,on='Cusip',how='outer')

Hilltop_Recent = Hilltop_Recent.fillna(0)
Hilltop_Recent.loc[Hilltop_Recent['Security']==0,'Security'] = Hilltop_Recent['Description_x']

"""
Calculate nessicary values

"""

Hilltop_Recent['Real_PNL_Change'] = Hilltop_Recent['Real PNL_x'] - Hilltop_Recent['Real PNL_y']
Hilltop_Recent['Real Discrepancy'] = Hilltop_Recent['Real_PNL_Change'] - Hilltop_Recent['P&L']
Hilltop_Recent['Quantity Change'] = Hilltop_Recent['Quantity_x'] - Hilltop_Recent['Quantity_y']

Hilltop_Individual_Book_Summary = Hilltop_Recent

Hilltop_Recent.rename(columns={'Quantity Change': 'HT Quantity Change',
                               'Quantity_x': 'HT New Quantity',
                               'Quantity_y': 'HT Old Quantity',
                               'Quantity_x_y': '-  HT Quantity  =',
                               'Real Discrepancy_y': 'TW vs. HT Real Discrepancy',
                               'Position': 'TW Quantity',
                               'Real PNL_x': 'HT New Real PNL',
                               'Account Name': 'Account',
                               'Quantity_y':'HT Old Quantity',
                               'Unreal PNL_x':'HT New Unreal PNL',
                               'Unreal PNL_y':'HT Old Unreal PNL',
                               'Real PNL_y':'HT Old Real PNL',
                               'P&L':'TW PNL',
                               'Quantity_y':'HT Old Quantity',
                               'Price_x':'Price'}, inplace=True)

# # Hilltop_Recent['TW Quantity'] = Hilltop_Recent['TW Quantity'] * 1000

Hilltop_Recent['HT Change in Quantity'] = Hilltop_Recent['HT New Quantity']-Hilltop_Recent['HT Old Quantity']
Hilltop_Recent['TW - HT Quantity Discrepancy'] = Hilltop_Recent['TW Quantity']-Hilltop_Recent['HT New Quantity']
Hilltop_Recent['Real PNL Change'] = Hilltop_Recent['HT New Real PNL']-Hilltop_Recent['HT Old Real PNL']
Hilltop_Recent['HT Real PNL Change'] = Hilltop_Recent['HT New Real PNL']-Hilltop_Recent['HT Old Real PNL']
Hilltop_Recent['Adj Unreal PNL Change'] = Hilltop_Recent['HT New Unreal PNL']-Hilltop_Recent['HT Old Unreal PNL']+Hilltop_Recent['HT Real PNL Change']
Hilltop_Recent['HT-TW PNL Discrepancy'] = Hilltop_Recent['HT Real PNL Change']-Hilltop_Recent['TW PNL']
Hilltop_Recent['Requirement Change'] = Hilltop_Recent['Requirement_x']-Hilltop_Recent['Requirement_y']
Hilltop_Recent['Filter Column'] = Hilltop_Recent['TW - HT Quantity Discrepancy'] + Hilltop_Recent['HT Change in Quantity'] + Hilltop_Recent['Adj Unreal PNL Change'] + Hilltop_Recent['HT-TW PNL Discrepancy'] + Hilltop_Recent['Requirement Change']
Hilltop_Recent = Hilltop_Recent[(Hilltop_Recent['Filter Column'] != 0)]

Hilltop_Recent = Hilltop_Recent[[
                                 'HT Quantity Change',                         
                                 'Security',                                       #A
                                 'Cusip',                                          #B
                                 'Account',                                        #C
                                 'Price',                                          #D
                                 'TW Quantity',                                    #E
                                 'HT New Quantity',                                #F
                                 'TW - HT Quantity Discrepancy',                   #G
                                 'HT Change in Quantity',                          #H
                                 'HT New Unreal PNL',                              #I
                                 'HT Old Unreal PNL',                              #J
                                 'Real PNL Change',                                #K
                                 'Adj Unreal PNL Change',                          #L
                                 'TW PNL',                                         #M
                                 'HT-TW PNL Discrepancy',                          #N
                                 'Requirement Change']]#9    

"""
# Drop Duplicated Values for Cusip and HT Quantity Change

"""
Hilltop_Recent = Hilltop_Recent.drop_duplicates(['Cusip', 'HT Quantity Change'])


"""
# Set up excel file naming path w/ today's date

"""
time = datetime.datetime.today()
current = time.strftime("%m.%d.%Y")
"""
# Write file to excel

"""
Hilltop_Recent.drop(
    [
        'HT Quantity Change'
    ],
    axis=1, inplace=True)


Hilltop_Individual_Summary = Hilltop_Recent
Hilltop_Recent.sort_values('TW PNL', axis=0, ascending=False, inplace=True)

filepath = 'C:/Users/ccraig/Desktop/New folder/'+ str(current) + 'PNL Discrepancy.xlsx'

writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
"""
            Summary Code

"""
Daily_Change_x = pd.read_excel(io=Recent, sheet_name='Summary')
Daily_Change_y = pd.read_excel(io=Old,sheet_name='Summary')
Daily_Change_x = Daily_Change_x[12:]
Daily_Change_y = Daily_Change_y[12:]
Daily_Change_x.rename(columns={'Account Number': 'Inventory Sub Totals',
                                'Total Available Funds':'Account Name',
                                'Unnamed: 2':'Position Type',
                                'Unnamed: 3':'Cost',
                                'Unnamed: 4':'Market Value',
                                'Unnamed: 5':'Unreal PNL',
                                'Unnamed: 6':'Requirement',
                                'Unnamed: 7':'Real PNL'}, inplace=True)
Daily_Change_x.reset_index(inplace = True)
Daily_Change_y.rename(columns={'Account Number': 'Inventory Sub Totals',
                                'Total Available Funds':'Account Name',
                                'Unnamed: 2':'Position Type',
                                'Unnamed: 3':'Cost',
                                'Unnamed: 4':'Market Value',
                                'Unnamed: 5':'Unreal PNL',
                                'Unnamed: 6':'Requirement',
                                'Unnamed: 7':'Real PNL'}, inplace=True)

Daily_Change_y.reset_index(inplace = True)
Daily_Change_x=pd.merge(Daily_Change_x, Daily_Change_y, on='index', how='left')
Daily_Change_x['Cost'] = Daily_Change_x['Cost_x']-Daily_Change_x['Cost_y']
Daily_Change_x['Market Value']=Daily_Change_x['Market Value_x']-Daily_Change_x['Market Value_y']
Daily_Change_x['Requirement']=Daily_Change_x['Requirement_x']-Daily_Change_x['Requirement_y']
Daily_Change_x['Unreal PNL']=Daily_Change_x['Unreal PNL_x']-Daily_Change_x['Unreal PNL_y']
Daily_Change_x['Real PNL']=Daily_Change_x['Real PNL_x']-Daily_Change_x['Real PNL_y']
Daily_Change_x = Daily_Change_x[['index','Account Name_x','Position Type_x','Cost','Market Value','Requirement','Unreal PNL','Real PNL']]
Daily_Change = Daily_Change_x.groupby('Account Name_x')['Cost','Market Value','Requirement', 'Unreal PNL','Real PNL'].sum()
Daily_Change['Account Name_x']=['K72 Muni Inv','K74 Corporates ','K76 S P Inv',
                                'K77 CD Inv','K78 Taxable Mun',
                                 'K79 Cali Tax Ex','K80 Muni Tax Ex',
                                'K81 Muni Tax','K82 Tax 0 Muni','L81 Sierra Comp',
                                  'M64 Sierra MBS','N90 CD','P01 Corp Floate','P02 Corp Sp','N88 Corp Notes']
Daily_Change = Daily_Change_x.append(Daily_Change, ignore_index = True,sort = False)
Daily_Change['Account Name_x'] = Daily_Change['Account Name_x'].replace({'K72 Muni Inv':'K72 Muni Inv Fl'})
Daily_Change.sort_values(['Account Name_x','Position Type_x'],inplace = True)
Daily_Change.rename(columns={'Account Name_x':'Account Name','Position Type_x':'Position Type'},inplace = True)
Daily_Change.drop('index',axis = 1, inplace = True)

"""
Create and Format Summary sheet

# """
Hilltop_Summary_new = pd.read_excel(io=Recent, sheet_name='Detail')
Hilltop_Summary_new = Hilltop_Summary_new.groupby(['Account Name']).agg({
                                                          'Unreal PNL':'sum',
                                                          'Real PNL':'sum',
                                                          'Requirement':'sum',
                                                          'Cost':'sum',
                                                          'Market Value':'sum'})
Hilltop_Summary_new = Hilltop_Summary_new[['Cost','Market Value','Requirement','Unreal PNL','Real PNL']]

Hilltop_Summary_new.loc['Column_Total'] = Hilltop_Summary_new.sum(numeric_only=True, axis=0)
Daily_Change.drop([1,5,7,9,11,13,15,17,19,21,33,35,36,37,38,39,40,41,42,43],inplace = True)
Daily_Change = Daily_Change.reindex([0,4,6,10,12,14,16,18,20,22,23,47,24,25,44,26,27,45,28,29,46,2,3,34,30,31,32])
Daily_Change.to_excel(writer,sheet_name ='Summary',index=False,startrow=20)
Hilltop_Summary_new.to_excel(writer,sheet_name ='Summary',index=True,startrow=1)

Hilltop_Recent_s['Change'] = Hilltop_Recent_s['Total Available Funds']-Hilltop_Old_s['Total Available Funds']
Hilltop_Recent_s = Hilltop_Recent_s[3:]
Hilltop_Recent_s = Hilltop_Recent_s[['Account Number','Total Available Funds','Change']]
Hilltop_Recent_s.to_excel(writer,sheet_name ='Summary',index=False,startrow=1,startcol=8)
workbook = writer.book

worksheet_summary = writer.sheets['Summary']
format1 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5'})
format2 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9',
                               'bg_color':'#f0f0f0'})
format3 = workbook.add_format()
format4 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5',
                               'bold': True,
                               'bg_color':'#f0f0f0'})
format5 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5',
                               'align':'center',
                               'bold': True,'bg_color':'#5551b8',
                               'font_color':'white'})
format7 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9',
                               'bg_color':'#f0f0f0'})
format7 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9'})
format8 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'11',
                               'bold': True})
format111 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'10',
                                'bg_color':'#f0f0f0',
                                'bold': True})
format_1 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'10',
                                'bg_color':'#f0f0f0'})
format_11 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'10',
                                'bold': True})
worksheet_summary.write('A2', 'Account Name',format111)
worksheet_summary.write('B2', 'Cost',format111)
worksheet_summary.write('C2', 'Market Value',format111)
worksheet_summary.write('D2', 'Requirment',format111)
worksheet_summary.write('E2','Unreal PNL',format111)
worksheet_summary.write('F2','Real PNL',format111)

worksheet_summary.write('I21', 'Account Name',format111)
worksheet_summary.write('J21', 'Cost',format111)
worksheet_summary.write('K21', 'Market Value',format111)
worksheet_summary.write('L21', 'Requirment',format111)
worksheet_summary.write('M21','Unreal PNL',format111)
worksheet_summary.write('N21','Real PNL',format111)

worksheet_summary.write('I2', 'Account Name',format111)
worksheet_summary.write('J2', 'Total Funds',format111)
worksheet_summary.write('K2', 'Change',format111)

worksheet_summary.write('A21', 'Account Name',format111)
worksheet_summary.write('B21', 'Position Type',format111)
worksheet_summary.write('C21', 'Cost',format111)
worksheet_summary.write('D21', 'Market Value',format111)
worksheet_summary.write('E21', 'Requirment',format111)
worksheet_summary.write('F21', 'Unreal PNL',format111)
worksheet_summary.write('G21', 'Real PNL',format111)

worksheet_summary.write('A18', 'Total',format5)
worksheet_summary.write('B46', 'Total Long',format5)
worksheet_summary.write('B47', 'Total Short',format5)
worksheet_summary.write('B48', 'Total',format5)

worksheet_summary.write('A3', 'K72 Muni Inv Fl',format_1)
worksheet_summary.write('A4', 'K74 Corporates',format_1)
worksheet_summary.write('A5', 'K76 S P Inv',format_1)
worksheet_summary.write('A6', 'K77 CD Inv',format_1)
worksheet_summary.write('A7','K78 Taxable Mun',format_1)
worksheet_summary.write('A8', 'K79 Cali Tax Ex',format_1)
worksheet_summary.write('A9', 'K80 Muni Tax Ex',format_1)
worksheet_summary.write('A10', 'K81 Muni Tax',format_1)
worksheet_summary.write('A11', 'K82 Tax 0 Muni',format_1)
worksheet_summary.write('A12','L81 Sierra Comp',format_1)
worksheet_summary.write('A13', 'M64 Sierra MBS',format_1)
worksheet_summary.write('A14', 'N90',format_1)
worksheet_summary.write('A15', 'P01 Corp Floate',format_1)
worksheet_summary.write('A16', 'P02 Corp SP',format_1)
worksheet_summary.write('A17', 'N88 Corp Notes',format_1)

worksheet_summary.write('A22', 'K72 Muni Inv Fl',format_1)
worksheet_summary.write('A23', 'K76 S P Inv',format_1)
worksheet_summary.write('A24', 'K77 CD Inv',format_1)
worksheet_summary.write('A25', 'K79 Cali Tax Ex',format_1)
worksheet_summary.write('A26', 'K80 Muni Tax Ex',format_1)
worksheet_summary.write('A27', 'K81 Muni Tax',format_1)
worksheet_summary.write('A28', 'K82 Tax 0 Muni',format_1)
worksheet_summary.write('A29', 'L81 Sierra Comp',format_1)
worksheet_summary.write('A30', 'M64 Sierra MBS',format_1)
worksheet_summary.write('A31', 'N88 Corp Notes',format_1)
worksheet_summary.write('A32', ' ',format_1)
worksheet_summary.write('A33', ' ',format_1)
worksheet_summary.write('A34', 'N90 CD ',format_1)
worksheet_summary.write('A35', ' ',format_1)
worksheet_summary.write('A36', ' ',format_1)
worksheet_summary.write('A37', 'P01 Corp Floate',format_1)
worksheet_summary.write('A38', ' ',format_1)
worksheet_summary.write('A39', ' ',format_1)
worksheet_summary.write('A40', 'P02 Corp SP',format_1)
worksheet_summary.write('A41', ' ',format_1)
worksheet_summary.write('A42', ' ',format_1)
worksheet_summary.write('A43', 'K74 Corp',format_1)
worksheet_summary.write('A44', ' ',format_1)
worksheet_summary.write('A45', ' ',format_1)

worksheet_summary.write('B22', 'Total',format_11)
worksheet_summary.write('B23', 'Total',format_11)
worksheet_summary.write('B24', 'Total',format_11)
worksheet_summary.write('B25', 'Total',format_11)
worksheet_summary.write('B26', 'Total',format_11)
worksheet_summary.write('B27', 'Total',format_11)
worksheet_summary.write('B28', 'Total',format_11)
worksheet_summary.write('B29', 'Total',format_11)
worksheet_summary.write('B30', 'Total',format_11)
worksheet_summary.write('B33', 'Total',format_11)
worksheet_summary.write('B36', 'Total',format_11)
worksheet_summary.write('B39', 'Total',format_11)
worksheet_summary.write('B42', 'Total',format_11)
worksheet_summary.write('B45', 'Total',format_11)


worksheet_summary.write('I3', 'Cost',format_1)
worksheet_summary.write('I4', 'Market Value',format_1)
worksheet_summary.write('I5', 'Unreal PNL',format_1)
worksheet_summary.write('I6', 'Real PNL',format_1)
worksheet_summary.write('I7', 'Requirement',format_1)
worksheet_summary.write('I8', 'Available Funds',format_1)
worksheet_summary.write('I9', 'Call/Excess',format_1)

worksheet_summary.set_column('A1:A14', 15) #2
worksheet_summary.set_column('B1:G14', 12, format1) #2

worksheet_summary.set_column('H1:H19',8,format1)
worksheet_summary.set_column('I1:I15',15,format1)
worksheet_summary.set_column('J1:N15',12,format1)

worksheet_summary.write('I22', 'Muni',format_1)
worksheet_summary.write_formula('J22','=SUM(C22:C23,C25:C29)')
worksheet_summary.write_formula('K22','=SUM(C22:C23,C25:C29)')
worksheet_summary.write_formula('L22','=SUM(C22:C23,C25:C29)')
worksheet_summary.write_formula('M22','=SUM(C22:C23,C25:C29)')
worksheet_summary.write_formula('N22','=SUM(C22:C23,C25:C29)')

worksheet_summary.write('I23', 'Corp',format_1)
worksheet_summary.write_formula('J23','=SUM(C33,C39,C42,C45)')
worksheet_summary.write_formula('K23','=SUM(C33,C39,C42,C45)')
worksheet_summary.write_formula('L23','=SUM(C33,C39,C42,C45)')
worksheet_summary.write_formula('M23','=SUM(C33,C39,C42,C45)')
worksheet_summary.write_formula('N23','=SUM(C33,C39,C42,C45)')

worksheet_summary.write('I24', 'CD',format_1)
worksheet_summary.write_formula('J24','=SUM(C36)')
worksheet_summary.write_formula('K24','=SUM(D36)')
worksheet_summary.write_formula('L24','=SUM(E36)')
worksheet_summary.write_formula('M24','=SUM(F36)')
worksheet_summary.write_formula('N24','=SUM(G36)')

worksheet_summary.write('I25', 'CMO',format_1)
worksheet_summary.write_formula('J25','=SUM(C30)')
worksheet_summary.write_formula('K25','=SUM(D30)')
worksheet_summary.write_formula('L25','=SUM(E30)')
worksheet_summary.write_formula('M25','=SUM(F30)')
worksheet_summary.write_formula('N25','=SUM(G30)')

worksheet_summary.set_zoom(95)

# worksheet_summary.set_column('L:P', None, None, {'hidden': True})

merge_format = workbook.add_format({
    'bold': 1,
    'border': 0,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': '#5551b8',
    'font_color':'white'})


worksheet_summary.merge_range('A1:F1', 'Monthly Summary', merge_format)
worksheet_summary.merge_range('A20:G20', 'Daily Change', merge_format)
worksheet_summary.merge_range('I20:N20', 'Product Summary',merge_format)
worksheet_summary.merge_range('I1:K1', 'Summary',merge_format)

"""
Create and format Detail sheet
"""

Hilltop_Recent.to_excel(writer, sheet_name='Detail', index=False)


worksheet = writer.sheets['Detail']

format1 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5',
                               'bottom':True})
format2 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9',
                               'bg_color':'#f0f0f0',
                               'bottom':True})
format3 = workbook.add_format()
format4 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5',
                               'bold': True,
                               'bottom':True,
                               'bg_color':'#f0f0f0'})
format5 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5',
                               'align':'center',
                               'bold': True,'bg_color':'#5551b8',
                               'font_color':'white',
                               'bottom':True})
format7 = workbook.add_format({'num_format': '#,##0',
                               'font_size':'9',
                               'bg_color':'#f0f0f0',
                               'bottom':True})
format6=workbook.add_format({'bottom':False})

format8=workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5',
                               'font_color':'green'})
format9=workbook.add_format({'num_format': '#,##0',
                               'font_size':'9.5',
                               'font_color':'red'})
worksheet_summary.conditional_format('K5:K6', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('K5:K6', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format9})
worksheet_summary.conditional_format('K8:K9', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('K8:K9', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format9})
worksheet_summary.conditional_format('F22:G65', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('F22:G65', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format9})

worksheet_summary.conditional_format('E3:F18', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('E3:F18', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format9})
worksheet_summary.conditional_format('M22:N27', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('M22:N27', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format9})
"""Requirement Change """
worksheet_summary.conditional_format('K7', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('K7', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format9})
worksheet_summary.conditional_format('E22:E65', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('E22:E65', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format9})

worksheet_summary.conditional_format('D3:D18', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('D3:D18', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format9})
worksheet_summary.conditional_format('L22:L27', {'type':'cell',
                                        'criteria': '<',
                                        'value':    0,
                                        'format':   format8})
worksheet_summary.conditional_format('L22:L27', {'type':'cell',
                                        'criteria': '>',
                                        'value':    0,
                                        'format':   format9})

worksheet_summary.insert_image('M1', 'P:/1. Individual Folders/Chad/Python Scripts/PNL Report/Logo.png',{'x_scale':0.5,'y_scale':0.5})

format1.set_text_wrap(True)
format3.set_bg_color('#cfcfcf')
format4.set_text_wrap(True)
format2.set_text_wrap(True)
format5.set_text_wrap(True)
format1.set_bottom(4)
format3.set_bottom(4)
format4.set_bottom(4)
format2.set_bottom(4)
format7.set_bottom(4)
# Format each colum to fit and display data correclty

worksheet.set_column('A:A', 24, format7) #2
worksheet.set_column('B:B', 9.5, format2) #2
worksheet.set_column('C:C', 14, format2) #3
worksheet.set_column('D:D', 11, format1)#4
worksheet.set_column('E:E', 11, format1)#5
worksheet.set_column('F:F', 11, format1)#6
worksheet.set_column('G:G', 11, format1)#7
worksheet.set_column('H:H', 11, format1)#8
worksheet.set_column('I:I', 11, format1)#9
worksheet.set_column('J:J', 11, format1)#10
worksheet.set_column('K:K', 11, format1)#11
worksheet.set_column('L:L', 11, format1)#12
worksheet.set_column('M:M', 11, format1)#13
worksheet.set_column('N:N', 11, format1)#14
worksheet.set_column('O:O', 11, format1)#15

worksheet.write('A1', 'Security',format5)
worksheet.write('B1', 'Cusip',format5)
worksheet.write('C1', 'Account',format5)
worksheet.write('D1', 'Price',format5)
worksheet.write('E1', 'TW QTY ',format5)
worksheet.write('F1', 'HT QTY',format5)
worksheet.write('G1', 'QTY Discrepancy',format5)
worksheet.write('H1', 'HT QTY Change',format5)
worksheet.write('I1', 'HT New Unreal PNL',format5)
worksheet.write('J1', 'HT Old Unreal PNL',format5)
worksheet.write('K1', 'Real PNL Change',format5)
worksheet.write('L1', 'Adj Unreal PNL Change',format5)
worksheet.write('M1', 'TW PNL',format5)
worksheet.write('N1', 'HT-TW PNL Discrep.',format5)
worksheet.write('O1', 'Requirement Change',format5)

worksheet.set_zoom(90)


QTY_Discrepancy = 'Discrepancy Between TW Quantity & Hilltop Quantity.'
HT_QTY_Change = 'Amount bought or sold on the day.'
Adj_Unreal_PNL_Change = 'Difference between Hilltop Unreal PNL and Real PNL. \n\nShows overall profit or loss while adjusting for the change in unrealized PNL. \n\nFor example: Bond has 1000 in Unreal PNL, it sells for a gain of 750. There would be a -250 change'
HT_TW_PNL_Discrepancy = 'Discrepancy between Hilltop  Real PNL and TW Real PNL.'
Requirement_Change = 'Increase or decrease in requirement.'

worksheet.write_comment('G1', QTY_Discrepancy, {'visible': 0})
worksheet.write_comment('I1', HT_QTY_Change, {'visible': 0})
worksheet.write_comment('O1', Adj_Unreal_PNL_Change, {'visible': 0,'x_scale': 1.5, 'y_scale': 1.8})
worksheet.write_comment('T1', HT_TW_PNL_Discrepancy, {'visible': 0})
worksheet.write_comment('V1', Requirement_Change, {'visible': 0})

worksheet.freeze_panes(1, 1)
worksheet.autofilter('A1:O20000')
worksheet.hide_gridlines(2)




Hilltop_x = Hilltop_Individual_Summary
Hilltop_y = Hilltop_x
# for cusip in Hilltop_Individual_Summary:
#     if Hilltop_Individual_Summary.loc['Security'] == 0:
#         Hilltop_Individual_Summary.loc['Cusip']['Security'] = Hilltop_Recent_x.loc['Cusip']['Description']

def Summary_Individual_Sheets(Hilltop_Individual_Summary):
    column_summary = ('TW - HT Quantity Discrepancy',
                      'HT Change in Quantity',
                      'Adj Unreal PNL Change',
                      'HT-TW PNL Discrepancy',
                      'Requirement Change')
    for item in column_summary:
        Hilltop_Individual_Summary[item] = Hilltop_Individual_Summary[item]
    Hilltop_QTY_DSP = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['TW - HT Quantity Discrepancy'] != 0)]
    Hilltop_HT_QTY_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['HT Change in Quantity'] != 0)]
    Hilltop_Adj_Unreal_PNL_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['Adj Unreal PNL Change'] != 0)]
    Hilltop_HT_TW_PNL_DSP = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['HT-TW PNL Discrepancy'] != 0)]
    Hilltop_Requirement_Change = Hilltop_Individual_Summary[(Hilltop_Individual_Summary['Requirement Change'] != 0)]
    return Hilltop_QTY_DSP,Hilltop_HT_QTY_Change,Hilltop_Adj_Unreal_PNL_Change, Hilltop_HT_TW_PNL_DSP,Hilltop_Requirement_Change

Individual_Sheets = Summary_Individual_Sheets(Hilltop_Individual_Summary)


Hilltop_QTY_DSP  = Individual_Sheets[0]
Hilltop_QTY_DSP = pd.merge(Hilltop_QTY_DSP, Hilltop_x, on='Cusip', how='left')
Hilltop_QTY_DSP['TW - HT Quantity Discrepancy_x'] = Hilltop_QTY_DSP['TW - HT Quantity Discrepancy_x'].abs()
Hilltop_QTY_DSP.sort_values('TW - HT Quantity Discrepancy_x', axis=0, ascending=False, inplace=True)
Hilltop_QTY_DSP = Hilltop_QTY_DSP[['Security_x','Cusip','Account_x','TW - HT Quantity Discrepancy_y']]
Hilltop_QTY_DSP.rename(columns={'Security_x': 'Security','Account_x':'Account','TW - HT Quantity Discrepancy_y':"QTY DSP"}, inplace=True)
QTY_DSP_Cleared_Positions_Drop = QTY_DSP_Cleared_Positions.drop('Position Notes',axis = 1)
Hilltop_QTY_DSP = Hilltop_QTY_DSP.append(QTY_DSP_Cleared_Positions_Drop)
Hilltop_QTY_DSP.drop_duplicates(subset ='Cusip',keep = False, inplace = True)
Hilltop_QTY_DSP = Hilltop_QTY_DSP[['Security','Account','Cusip','QTY DSP']]
Hilltop_QTY_DSP.to_excel(writer, sheet_name = 'HT QTY DSP', index=False)
worksheet_Hilltop_QTY_DSP = writer.sheets['HT QTY DSP']
QTY_DSP_Cleared_Positions= QTY_DSP_Cleared_Positions[['Security','Account','Cusip','QTY DSP','Position Notes']]
QTY_DSP_Cleared_Positions.to_excel(writer,sheet_name ='HT QTY DSP',index=False,startrow=1,startcol=6)

Hilltop_HT_TW_PNL_DSP = Individual_Sheets[3]
Hilltop_HT_TW_PNL_DSP.sort_values('HT-TW PNL Discrepancy', axis=0, ascending=False, inplace=True)
Hilltop_HT_TW_PNL_DSP = pd.merge(Hilltop_HT_TW_PNL_DSP, Hilltop_x, on='Cusip', how='left')
Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_x'] = Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_x'].abs()
Hilltop_HT_TW_PNL_DSP.sort_values('HT-TW PNL Discrepancy_x', axis=0, ascending=False, inplace=True)
Hilltop_HT_TW_PNL_DSP = Hilltop_HT_TW_PNL_DSP[['Security_x','Cusip','Account_x','HT-TW PNL Discrepancy_y']]
Hilltop_HT_TW_PNL_DSP = Hilltop_HT_TW_PNL_DSP[(Hilltop_HT_TW_PNL_DSP['HT-TW PNL Discrepancy_y'] > 10)]
Hilltop_HT_TW_PNL_DSP.to_excel(writer, sheet_name = 'HT-TW PNL DSP', index=False)
worksheet_Hilltop_HT_TW_PNL_DSP = writer.sheets['HT-TW PNL DSP']


Hilltop_Adj_Unreal_PNL_Change = Individual_Sheets[2]
Hilltop_Adj_Unreal_PNL_Change = pd.merge(Hilltop_Adj_Unreal_PNL_Change, Hilltop_x, on='Cusip', how='left')
Hilltop_Adj_Unreal_PNL_Change['Adj Unreal PNL Change_x'] = Hilltop_Adj_Unreal_PNL_Change['Adj Unreal PNL Change_x'].abs()
Hilltop_Adj_Unreal_PNL_Change.sort_values('Adj Unreal PNL Change_x', axis=0, ascending=False, inplace=True)
Hilltop_Adj_Unreal_PNL_Change = Hilltop_Adj_Unreal_PNL_Change[['Security_x','Cusip','Account_x','Adj Unreal PNL Change_y']]
Hilltop_Adj_Unreal_PNL_Change.to_excel(writer, sheet_name = 'Adj Unreal PNL Change', index=False)
worksheet_Hilltop_Adj_Unreal_PNL_Change = writer.sheets['Adj Unreal PNL Change']


Hilltop_Requirement_Change = Individual_Sheets[4]
Hilltop_Requirement_Change = pd.merge(Hilltop_Requirement_Change, Hilltop_x, on='Cusip', how='left')
Hilltop_Requirement_Change = Hilltop_Requirement_Change[['Cusip','Security_x','Account_x','Requirement Change_x','Requirement Change_y']]
Hilltop_Requirement_Change['Requirement Change_x'] = Hilltop_Requirement_Change['Requirement Change_x'].abs()
Hilltop_Requirement_Change.sort_values('Requirement Change_x', axis=0, ascending=False, inplace=True)
Hilltop_Requirement_Change = Hilltop_Requirement_Change[['Security_x','Cusip','Account_x','Requirement Change_y']]
Hilltop_Requirement_Change.to_excel(writer, sheet_name = 'Requirement Change', index=False)
worksheet_Hilltop_Requirement_Change = writer.sheets['Requirement Change']

worksheet_Hilltop_QTY_DSP.set_column('A:A', 20, format7) #2
worksheet_Hilltop_QTY_DSP.set_column('B:B', 12, format2) #2
worksheet_Hilltop_QTY_DSP.set_column('C:C', 12, format7) #3
worksheet_Hilltop_QTY_DSP.set_column('D:D', 12, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('D:D', 12, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('E:E', 30, format1)

worksheet_Hilltop_QTY_DSP.write('A1', 'Security',format5)
worksheet_Hilltop_QTY_DSP.write('B1', 'Account',format5)
worksheet_Hilltop_QTY_DSP.write('C1', 'Cusip',format5)
worksheet_Hilltop_QTY_DSP.write('D1', 'QTY DSP',format5)
worksheet_Hilltop_QTY_DSP.write('E1', 'Position Notes',format5)

worksheet_Hilltop_QTY_DSP.write('G2', 'Security',format5)
worksheet_Hilltop_QTY_DSP.write('H2', 'Account',format5)
worksheet_Hilltop_QTY_DSP.write('I2', 'Cusip',format5)
worksheet_Hilltop_QTY_DSP.write('J2', 'QTY DSP',format5)
worksheet_Hilltop_QTY_DSP.write('K2', 'Position Notes',format5)

worksheet_Hilltop_QTY_DSP.merge_range('G1:K1', 'Cleared QTY DSP',merge_format)

worksheet_Hilltop_QTY_DSP.set_column('G:G', 20, format1) #2
worksheet_Hilltop_QTY_DSP.set_column('H:H', 12, format1) #2
worksheet_Hilltop_QTY_DSP.set_column('I:I', 12, format1) #3
worksheet_Hilltop_QTY_DSP.set_column('J:J', 15, format1)#4
worksheet_Hilltop_QTY_DSP.set_column('K:K', 30, format1)#4

worksheet_Hilltop_QTY_DSP.freeze_panes(1, 1)
worksheet_Hilltop_QTY_DSP.autofilter('A1:V20000')
worksheet_Hilltop_QTY_DSP.hide_gridlines(2)
worksheet_Hilltop_QTY_DSP.protect('welcome123')
worksheet_Hilltop_QTY_DSP.set_zoom(90)
"""
"""
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('A:A', 30, format7) #2
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('B:B', 12, format2) #2
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('C:C', 12, format7) #3
worksheet_Hilltop_Adj_Unreal_PNL_Change.set_column('D:D', 12, format1)#4

worksheet_Hilltop_Adj_Unreal_PNL_Change.write('A1', 'Security',format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('B1', 'Cusip',format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('C1', 'Account',format5)
worksheet_Hilltop_Adj_Unreal_PNL_Change.write('D1', 'Adj Unreal PNL Change ',format5)
"""
"""
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('A:A', 30, format7) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('B:B',  9, format2) #2
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('C:C', 12, format7) #3
worksheet_Hilltop_HT_TW_PNL_DSP.set_column('D:D', 11, format1)#4

worksheet_Hilltop_HT_TW_PNL_DSP.write('A1', 'Security',format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('B1', 'Cusip',format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('C1', 'Account',format5)
worksheet_Hilltop_HT_TW_PNL_DSP.write('D1', 'PNL DSP ',format5)

"""
"""
worksheet_Hilltop_Requirement_Change.set_column('A:A', 30, format7) #2
worksheet_Hilltop_Requirement_Change.set_column('B:B', 12, format2) #2
worksheet_Hilltop_Requirement_Change.set_column('C:C', 12, format7) #3
worksheet_Hilltop_Requirement_Change.set_column('D:D', 12, format1)#4

worksheet_Hilltop_Requirement_Change.write('A1', 'Security',format5)
worksheet_Hilltop_Requirement_Change.write('B1', 'Cusip',format5)
worksheet_Hilltop_Requirement_Change.write('C1', 'Account',format5)
worksheet_Hilltop_Requirement_Change.write('D1', 'Requirement Change',format5)



worksheet_Hilltop_QTY_DSP.freeze_panes(1, 1)
worksheet_Hilltop_QTY_DSP.autofilter('A1:E20000')
worksheet_Hilltop_QTY_DSP.hide_gridlines(2)

worksheet_Hilltop_HT_TW_PNL_DSP.freeze_panes(1, 1)
worksheet_Hilltop_HT_TW_PNL_DSP.autofilter('A1:D20000')
worksheet_Hilltop_HT_TW_PNL_DSP.hide_gridlines(2)


worksheet_Hilltop_Adj_Unreal_PNL_Change.freeze_panes(1, 1)
worksheet_Hilltop_Adj_Unreal_PNL_Change.autofilter('A1:D20000')
worksheet_Hilltop_Adj_Unreal_PNL_Change.hide_gridlines(2)

worksheet_Hilltop_Requirement_Change.freeze_panes(1, 1)
worksheet_Hilltop_Requirement_Change.autofilter('A1:D20000')
worksheet_Hilltop_Requirement_Change.hide_gridlines(2)

worksheet_summary.hide_gridlines(2)

workbook.close()

import win32com.client
from win32com.client import Dispatch, constants
const=win32com.client.constants

olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = "PNL Report"
text_body = 'PNL Report for '+today+' & '+yesterday+'.'
newMail.To = 'ccraig90@bloomberg.net'#; lankowsky@bloomberg.net; jdean18@bloomberg.net; jblamire@bloomberg.net'
newMail.body = text_body
newMail.Attachments.Add('C:/Users/ccraig/Desktop/New folder/'+ str(current) + str('PNL Discrepancy.xlsx'))
newMail.display()
newMail.Send()

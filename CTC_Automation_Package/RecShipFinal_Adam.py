import numpy as np
import pandas as pd
import datetime as dt
from sqlalchemy.engine import create_engine

import cx_Oracle
print(cx_Oracle.version)
#check the Oracle Version using: select * from v$version
import sys
print(sys.executable)
print(sys.version)
print(sys.version_info)
#follow this link to solove issue: cx_Oracle error. DPI-1047: Cannot locate a 64-bit Oracle Client library
# https://stackoverflow.com/questions/56119490/cx-oracle-error-dpi-1047-cannot-locate-a-64-bit-oracle-client-library?answertab=trending#tab-top
import os
import platform
# This is the path to the ORACLE client files
lib_dir = r"C:\Users\yongpeng.fu\OneDrive - Canadian Tire\Desktop\CTC Work\From Olivia Lee\CTC_June202022_2\instantclient_21_6"

# Diagnostic output to verify 64 bit arch and list files
print("ARCH:", platform.architecture())
print("FILES AT lib_dir:")
name_string = ""
for name in os.listdir(lib_dir):
    name_string = name_string + name + " "
print(name_string)

#you just need to run the following snipet for once, thats why I comment them out since I have ran this    
try:
    cx_Oracle.init_oracle_client(lib_dir=lib_dir)
except Exception as err:
    print("Error connecting: cx_Oracle.init_oracle_client()")
    print(err);
    sys.exit(1);

cx_Oracle.clientversion()

DIALECT = 'oracle'
SQL_DRIVER = 'cx_oracle'
USERNAME = "ablaine" #enter your username
PASSWORD = "fUtk!n8MgYGSKUnJ" #enter your password
HOST = "p9cpwpjdadb01" #enter the oracle db host url
PORT = 25959 # enter the oracle port number
SERVICE = "FR01PR" # enter the oracle db service name
ENGINE_PATH_WIN_AUTH = DIALECT + '+' + SQL_DRIVER + '://' + USERNAME + ':' + PASSWORD +'@' + HOST + ':' + str(PORT) + '/?service_name=' + SERVICE

engine = create_engine(ENGINE_PATH_WIN_AUTH)

db = pd.read_sql_query('''
    Select i.U_CATEGORY,i.U_COMMODITY, i.u_style, i.u_choice, i.U_SPECIFICCOLORNAME, i.U_SIZE_ONE, R.ITEM, r.source, r.dest, r.QTY, R.ORDERPLACEDATE,r.SCHEDARRIVDATE, s.OH ,dmd.u_permretailprice retail
    , (SELECT SUM(nvl(f.qty,0)) FROM scpomgr.cds_fcstview f where f.dmdunit = s.item and f.loc = s.loc and f.startdate between sysdate - 7 and sysdate +42) "total fcst"
    from SCPOMGR.recship r
    inner join SCPOMGR.item i on i.item = r.item
    INNER JOIN SCPOMGR.SKU S ON S.ITEM = r.ITEM AND S.LOC = r.dest
    inner join SCPOMGR.dmdunit dmd on i.item = dmd.dmdunit
    where
    i.u_style in ('S_1200BZ',
    'S_28116A',
    'S_28117A',
    'S_28118A',
    'S_28154A',
    'S_28155A',
    'S_67181',
    'S_67182',
    'S_67479-HH',
    'S_67482',
    'S_6ARCDHASCBOGOF1',
    'S_6ARCDHASCBOGOF9',
    'S_6ARDHHAS1064053',
    'S_6AREDHASHBSPTS2',
    'S_75729',
    'S_79645',
    'S_79646'
    )
    and R.source IN ('000155')--'V100033'
    --and r.dest in ('000357','000534','000536','000540') 
    and R.ORDERPLACEDATE between sysdate - 0 and sysdate +31
    ''', engine)

    #dont forget to close the engine to release the resource you have used on the server.
engine.dispose()

if dt.date.today().weekday() == 4:
    OneDay = dt.date.today() + dt.timedelta(days=3) 
else:
    OneDay = dt.date.today() + dt.timedelta(days=1)

db.columns = db.columns.str.upper()

db.sort_values(by=['DEST',  'ITEM','ORDERPLACEDATE'], ascending = [True, True, True], inplace=True )
db = db.reset_index(drop=True)

db['GAP'] = db.groupby(['ITEM','DEST' ])['ORDERPLACEDATE'].diff()
db['AVGQTY'] = db.groupby(['ITEM'])['QTY'].transform('mean')
db['STDQTY'] = db.groupby(['ITEM'])['QTY'].transform('std')
db['ZSCORE'] = (db['QTY']-db['AVGQTY']) / (db['STDQTY']+0.000001)

db['QTY'] = db['QTY'].astype('int64')
db['OH'] = db['OH'].astype('int64')
db['total fcst'] = db['total fcst'].astype('int64')

db['GAP']= db['GAP'].astype('timedelta64[D]')
db['GAP'] = db['GAP'].fillna(0)

db['SUPPLY'] =  round(db['OH'] / ((db['total fcst']+0.0001) / 7) , 2)

db['STATUS'] = 'Original'

db['2W Prior'] = np.where((db.loc[:,'GAP']< 14) & (db.loc[:,'GAP'] > 0), True, False)
db['1W Prior'] = np.where((db['GAP']< 8) & (db.loc[:,'GAP'] > 0), True, False)
db['2W lowQ'] = np.where((db['GAP']< 14) & (db['ZSCORE'] < -.15) & (db.loc[:,'GAP'] > 0), True, False)
db['1W lowQ'] = np.where((db['GAP']< 8) & (db['ZSCORE'] < -.15) & (db.loc[:,'GAP'] > 0), True, False)

orig = db.copy()
db['GAP']= db['GAP'].astype('timedelta64[D]')
db['ORDERPLACEDATE'] = db['ORDERPLACEDATE'].dt.date

OrigQ = db[db['ORDERPLACEDATE'] == OneDay].QTY.sum()

db1 = db.copy()
for i in range (1, len(db1)):
    if (db1.loc[i, '1W lowQ'] == True) & (db1.loc[i - 1, '1W lowQ'] == True):
        db1.loc[i, '1W lowQ'] = False 

for i in range (1, len(db1)):
    if (db1.loc[i, '1W lowQ'] == True) :
        if db1.loc[i, 'SUPPLY'] < 10 :
            db1.loc[i-1 , 'QTY'] = db1.loc[i-1 , 'QTY'] + db1.loc[i, 'QTY']
            db1.loc[i , 'QTY'] = 0
            db1.loc[i-1, 'STATUS'] = 'Promoted'
            db1.loc[i, 'STATUS'] = 'Removed'

        else:
            db1.loc[i, 'QTY'] = db1.loc[i-1 , 'QTY'] + db1.loc[i, 'QTY']
            db1.loc[i-1 , 'QTY'] = 0
            db1.loc[i, 'STATUS'] = 'Postponed'
            db1.loc[i-1, 'STATUS'] = 'Removed'

for i in range (1, len(db1)):
    if (db1.loc[i, '2W Prior'] == True) & (db1.loc[i, 'SUPPLY'] < 3) & (db1.loc[i, 'STATUS'] == 'Original') :
        db1.loc[i - 1, 'QTY'] =  db1.loc[i-1 , 'QTY'] + db1.loc[i, 'QTY']
        db1.loc[i , 'QTY'] = 0
        db1.loc[i-1, 'STATUS'] = 'Prioritized'
        db1.loc[i, 'STATUS'] = 'Removed'

db2 = db1[db1['ORDERPLACEDATE'] == OneDay]
promote = db2.STATUS.str.count("Promoted").sum()
postpone = db2.STATUS.str.count("Removed").sum()
priority = db2.STATUS.str.count("Prioritized").sum()
db2.loc[:,'ORDERPLACEDATE'] = pd.to_datetime(db2.loc[:,'ORDERPLACEDATE'])


final1d = db2.copy()
final1d = final1d[['ITEM','SOURCE','DEST', 'ITEM', 'ORDERPLACEDATE', 'ORDERPLACEDATE', 'ORDERPLACEDATE', 'ORDERPLACEDATE', 'QTY']].copy()
final1d.drop(final1d.index[final1d['QTY'] == 0], inplace =True)
final1d.insert(3, 'Type', 2)
final1d.insert(4, 'SeqNum', range(1000, 1000 + len(final1d)))
final1d['U_approvedSW'] = 1
final1d.columns = ['Sku', 'Source', 'Destination', 'Type', 'SeqNum', 'Primary Item', 'AvailtoShipDat', 'DepartureDate', 'NeedArrivDate', 'ShedArriveDate', 'Qty', 'U_ApprovedSW']
final1d.drop(final1d.index[final1d['Qty'] == 0], inplace =True)
final1d['NeedArrivDate'] = final1d['NeedArrivDate'] + dt.timedelta(days = 7)
final1d['ShedArriveDate'] = final1d['ShedArriveDate'] + dt.timedelta(days = 7)

final1d['AvailtoShipDat'] = final1d['AvailtoShipDat'].dt.strftime("%m/%d/%Y")
final1d['DepartureDate'] = final1d['DepartureDate'].dt.strftime("%m/%d/%Y")
final1d['NeedArrivDate'] = final1d['NeedArrivDate'].dt.strftime("%m/%d/%Y")
final1d['ShedArriveDate'] = final1d['ShedArriveDate'].dt.strftime("%m/%d/%Y")

HiQ = db1[db1['ZSCORE'] > 5]

NewQ = final1d.Qty.sum()

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Revised_RecShip2.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
final1d.to_excel(writer, sheet_name = OneDay.strftime("%m-%d-%Y"))
HiQ.to_excel(writer, sheet_name='High Quantity')
db1.to_excel(writer, sheet_name='Updated')
orig.to_excel(writer, sheet_name='Original')

workbook  = writer.book
summary = workbook.add_worksheet('Summary')

text = "The function postponed {} RecShips scheduled for tomorrow. {} RecShips were promoted from next week. {} were prioritized because of low supply. The total daily Qty went from {} to {}".format(postpone, promote, priority, OrigQ, NewQ)
summary.insert_textbox(0, 0, text)

workbook.close()


# Close the Pandas Excel writer and output the Excel file.
writer.save()
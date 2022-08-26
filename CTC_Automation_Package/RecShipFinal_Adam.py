import os
import subprocess
import sys
import pkg_resources


#check if you have all the package, if not install it
def recShip_prep():
    def install(package):
        if (package not in {pkg.key for pkg in pkg_resources.working_set}):
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
    for pkg in ['cx-oracle', 'cx-oracle', 'sqlalchemy', 'pandas', 'numpy', 'datetime', 'wget', 'zipp']:
        install(pkg)
    import wget
    from zipfile import ZipFile

    # check if you have instantclient-basic-windows.x64-21.6.0.0.0dbru, if you have run initiation
    # if not then you download and run the initiation
    download_link = 'https://download.oracle.com/otn_software/nt/instantclient/216000/instantclient-basic-windows.x64-21.6.0.0.0dbru.zip'
    if not os.path.exists('./instantclient_21_6'):
        wget.download(download_link)
        with ZipFile('./instantclient-basic-windows.x64-21.6.0.0.0dbru.zip', 'r') as zipObj:
            # Extract all the contents of zip file in current directory
            zipObj.extractall()


def recShip_output(style_list, kargs):
    import cx_Oracle
    from sqlalchemy.engine import create_engine
    import pandas as pd
    import numpy as np
    import datetime as dt

    #prepare 
    lib_dir = './instantclient_21_6'
    name_string = ""
    for name in os.listdir(lib_dir):
        name_string = name_string + name + " "
    #you just need to run the following snipet for once.
    try:
        cx_Oracle.init_oracle_client(lib_dir=lib_dir)
    except:
        pass

    #Connect to the Oracle database using SQLAIchemy (recommended)
    DIALECT = kargs['DIALECT'] #'oracle'
    SQL_DRIVER = kargs['SQL_DRIVER']#cx_oracle'
    USERNAME = kargs['USERNAME']#"ablaine" #enter your username
    PASSWORD = kargs['PASSWORD']#"fUtk!n8MgYGSKUnJ" #enter your password
    HOST = kargs['HOST']#"p9cpwpjdadb01" #enter the oracle db host url
    PORT = kargs['PORT']#25959 # enter the oracle port number
    SERVICE = kargs['SERVICE']#"FR01PR" # enter the oracle db service name
    ENGINE_PATH_WIN_AUTH = DIALECT + '+' + SQL_DRIVER + '://' + USERNAME + ':' + PASSWORD +'@' + HOST + ':' + str(PORT) + '/?service_name=' + SERVICE
    engine = create_engine(ENGINE_PATH_WIN_AUTH)

    db = pd.read_sql_query(f'''
    Select i.U_CATEGORY,i.U_COMMODITY, i.u_style, i.u_choice, i.U_SPECIFICCOLORNAME, i.U_SIZE_ONE, R.ITEM, r.source, r.dest, r.QTY, R.ORDERPLACEDATE,r.SCHEDARRIVDATE, s.OH ,dmd.u_permretailprice retail
    , (SELECT SUM(nvl(f.qty,0)) FROM scpomgr.cds_fcstview f where f.dmdunit = s.item and f.loc = s.loc and f.startdate between sysdate - 7 and sysdate +42) "total fcst"
    from SCPOMGR.recship r
    inner join SCPOMGR.item i on i.item = r.item
    INNER JOIN SCPOMGR.SKU S ON S.ITEM = r.ITEM AND S.LOC = r.dest
    inner join SCPOMGR.dmdunit dmd on i.item = dmd.dmdunit
    where
    i.u_style in ({','.join(map(str, style_list))})
    and R.source IN ('000155')--'V100033'
    --and r.dest in ('000357','000534','000536','000540') 
    and R.ORDERPLACEDATE between sysdate - 0 and sysdate +31
    ''', engine)

    #dont forget to close the engine to release the resource you have used on the server.
    # engine.dispose()

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
    db2['ORDERPLACEDATE'] = pd.to_datetime(db2['ORDERPLACEDATE'])


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
    writer = pd.ExcelWriter(os.path.expanduser("~\\OneDrive - Canadian Tire\\Desktop\\Revised_RecShip2_" + os.getlogin() + ".xlsx"), engine='xlsxwriter')

    # Write each dataframe to a different worksheet.
    final1d.to_excel(writer, sheet_name = OneDay.strftime("%m-%d-%Y"))
    HiQ.to_excel(writer, sheet_name='High Quantity')
    db1.to_excel(writer, sheet_name='Updated')
    orig.to_excel(writer, sheet_name='Original')

    workbook  = writer.book
    summary = workbook.add_worksheet('Summary')

    text = "The function postponed {} RecShips scheduled for tomorrow. {} RecShips were promoted from next week. {} were prioritized because of low supply. The total daily Qty went from {} to {}".format(postpone, promote, priority, OrigQ, NewQ)
    summary.insert_textbox(0, 0, text)


    # Close the Pandas Excel writer and output the Excel file.
    writer.save()
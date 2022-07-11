# July 3 2022: Build bot for automating Selling Curve Builder

#This is the information procided by Leszek from VB: 
# Function OracleConnection(userid As String, password As String) As String  
#     OracleConnection = "Provider=msdaora;Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=p9icwpmpmmdb01.idm.ad.ctc (PORT=26001))(CONNECT_DATA=(SERVICE_NAME=mpmm01pr)));User Id=" & userid & ";Password=" & password & ";"  
#     End Function

import cx_Oracle
print(cx_Oracle.version)
#check the Oracle Version using: select * from v$version
import os #get access to the operation system
import sys
print(sys.executable)
print(sys.version)
print(sys.version_info)
#follow this link to solove issue: cx_Oracle error. DPI-1047: Cannot locate a 64-bit Oracle Client library
# https://stackoverflow.com/questions/56119490/cx-oracle-error-dpi-1047-cannot-locate-a-64-bit-oracle-client-library?answertab=trending#tab-top
import os
import platform
# This is the path to the ORACLE client files
lib_dir = r"..\instantclient-basic-windows.x64-21.6.0.0.0dbru\instantclient_21_6"

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


#Connect to the Oracle database using SQLAIchemy (recommended)
from sqlalchemy.engine import create_engine

DIALECT = 'oracle'
SQL_DRIVER = 'cx_oracle'
USERNAME = "ypfu" #enter your username
PASSWORD = "cT4K3tgta8NnS!QK" #enter your password
HOST = "p9cpwpjdadb01" #enter the oracle db host url
PORT = 25959 # enter the oracle port number
SERVICE = "FR01PR" # enter the oracle db service name
ENGINE_PATH_WIN_AUTH = DIALECT + '+' + SQL_DRIVER + '://' + USERNAME + ':' + PASSWORD +'@' + HOST + ':' + str(PORT) + '/?service_name=' + SERVICE

engine = create_engine(ENGINE_PATH_WIN_AUTH)

## Load the XX XX_HIST MOD Template
# The xlsx is stored at K:\Logistics\_DFP Reports - Menswear\2022\F22 Setup
# NOTE: sometimes the template has missing value for K column for example, you want to double chcck if the plate excel sheet is ok before you run the following code.

import pandas as pd
import numpy as np
from datetime import date
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
#Read the excel file and get a list of sheets. Then chose and load the sheets.
HIST_MOD_Template = pd.ExcelFile(file_path)
print(HIST_MOD_Template.sheet_names)
#to read just 'TRACKER' sheet to dataframe
HIST_MOD_Template_TRACKER = pd.read_excel(HIST_MOD_Template, sheet_name = "TRACKER")
HIST_MOD_Template_TRACKER.dropna(how = "all", inplace = True, subset=["Style Number"])
#create the Curve Name and Copy it into Curve ID column as well
HIST_MOD_Template_TRACKER['CURVE NAME'] = \
pd.Series(np.where(HIST_MOD_Template_TRACKER["HIST MOD SETTING - DISSECTION"].isnull(), 
         HIST_MOD_Template_TRACKER["HIST MOD SETTING - COMMODITY"].str.split().str[0],
        HIST_MOD_Template_TRACKER["HIST MOD SETTING - DISSECTION"].str.split().str[0])) \
+ "-" + HIST_MOD_Template_TRACKER["Selling Cycle"]\
+ "-" + HIST_MOD_Template_TRACKER["START WK"].str.replace(' ','')\
+ "-" + HIST_MOD_Template_TRACKER["END WK"].str.replace(' ','')
HIST_MOD_Template_TRACKER['CURVE_ID'] = HIST_MOD_Template_TRACKER['CURVE NAME']
HIST_MOD_Template_TRACKER['SUBMIT_DATE'] = pd.Timestamp("today").strftime('%Y/%m/%d')

#connetion string using SQLAIchemy
DIALECT = 'oracle'
SQL_DRIVER = 'cx_oracle'
USERNAME = "ypfu" #enter your username
PASSWORD = "cT4K3tgta8NnS!QK" #enter your password
HOST = "p9cpwpjdadb01" #enter the oracle db host url
PORT = 25959 # enter the oracle port number
SERVICE = "FR01PR" # enter the oracle db service name
ENGINE_PATH_WIN_AUTH = DIALECT + '+' + SQL_DRIVER + '://' + USERNAME + ':' + PASSWORD +'@' + HOST + ':' + str(PORT) + '/?service_name=' + SERVICE
engine = create_engine(ENGINE_PATH_WIN_AUTH)

#scan through each row of the dataframe to retrive data from the database
#create a function to compare "START WK" and "END WK" for each row
def wk_list(wk1, wk2):
    wk1 = int(wk1.split()[1])
    wk2 = int(wk2.split()[1])
    if wk2 > wk1:
        return np.arange(wk1,wk2+1)
    else:
        return np.append(np.arange(wk1,53),np.arange(1,wk2+1))
    
#map the selling cycle to the correct notation as follows:
def selling_cycle_mapper(cycle):
    return {'All Season':'AS-All Season',
           "Winter Basics":"WB-Winter Basics",
           "Fall Basics":"FB-Fall Basics",
           "Spring":"SP-Spring",
           "Fall":"FA-Fall",
           "SB-Spring Basics":"SB-Spring Basics",
           "Holiday":"FH-Holiday"}[cycle]
    
#create a new array with 0 indicate to use COMMODITY and 1 for DISSECTION during SQL query
comm_or_diss = np.column_stack((np.where(HIST_MOD_Template_TRACKER["HIST MOD SETTING - DISSECTION"].isnull(), 
         "M_" + HIST_MOD_Template_TRACKER["HIST MOD SETTING - COMMODITY"].str.split().str[0],
        HIST_MOD_Template_TRACKER["HIST MOD SETTING - DISSECTION"].str.split().str[0]),
                              np.where(HIST_MOD_Template_TRACKER["HIST MOD SETTING - DISSECTION"].isnull(), 
                                       0,1), HIST_MOD_Template_TRACKER[["START WK", "END WK", "CURVE NAME", "Selling Cycle"]].to_numpy()))

#dont forget to close the engine to release the resource you have used on the server.
engine.dispose()
#define query to pull data:
selling_curve_query_comm = '''
Select c.CALCWY WEEK
,sum((Case When c.CALCYR = 2019 Then nvl(h.qty,0) Else 0 End)) Year_2019
,sum((Case When c.CALCYR = 2020 Then nvl(h.qty,0) Else 0 End)) Year_2020
,sum((Case When c.CALCYR = 2021 Then nvl(h.qty,0) Else 0 End)) Year_2021
From scpomgr.cds_histview h
Inner Join DFREPORTING.CALDTLEE c on h.startdate = c.CALCUR
--Where h.dmdunit in (Select distinct dmdunit from scpomgr.dmdunit where u_style in ('S_2DIADHAS-419') and dmdunit like 'C_%')
--Where h.dmdunit in (Select distinct dmdunit from scpomgr.dmdunit where u_dissection = :diss and dmdunit like 'C_%')
Where h.dmdunit in (Select distinct dmdunit from scpomgr.dmdunit where u_commodity = :comm and dmdunit like 'C_%')
And h.loc = 'ALL'
And h.event = 'TOTAL'
And c.CALCYR between 2019 and 2021
Group By c.CALCWY
Order By c.CALCWY
'''
selling_curve_query_diss = '''
Select c.CALCWY WEEK
,sum((Case When c.CALCYR = 2019 Then nvl(h.qty,0) Else 0 End)) Year_2019
,sum((Case When c.CALCYR = 2020 Then nvl(h.qty,0) Else 0 End)) Year_2020
,sum((Case When c.CALCYR = 2021 Then nvl(h.qty,0) Else 0 End)) Year_2021
From scpomgr.cds_histview h
Inner Join DFREPORTING.CALDTLEE c on h.startdate = c.CALCUR
--Where h.dmdunit in (Select distinct dmdunit from scpomgr.dmdunit where u_style in ('S_2DIADHAS-419') and dmdunit like 'C_%')
Where h.dmdunit in (Select distinct dmdunit from scpomgr.dmdunit where u_dissection = :diss and dmdunit like 'C_%')
--Where h.dmdunit in (Select distinct dmdunit from scpomgr.dmdunit where u_commodity = :comm and dmdunit like 'C_%')
And h.loc = 'ALL'
And h.event = 'TOTAL'
And c.CALCYR between 2019 and 2021
Group By c.CALCWY
Order By c.CALCWY
'''

#define a function to decide which query to use for the input weeks
def query_to_use(holder, indicator, wk1, wk2):
    if indicator:
        returned_result = pd.read_sql_query(selling_curve_query_diss, engine, params = [holder])
    else:
        returned_result = pd.read_sql_query(selling_curve_query_comm, engine, params = [holder])
    #remove the last row because we are only considering 52 weeks
    returned_result = returned_result[:-1]
    returned_result["Avg2019-2021"] = returned_result[["year_2019", "year_2021"]].mean(axis = 1).apply(np.ceil).astype(int)
    returned_result.loc[returned_result["week"].isin(np.setdiff1d(np.arange(1,53),wk_list(wk1, wk2))), "Avg2019-2021"] = 0
    returned_result["Avg2019-2021ratio"] = (returned_result["Avg2019-2021"]/returned_result["Avg2019-2021"].sum())
    returned_result.index = np.arange(1,len(returned_result) + 1)
    return returned_result

#scan through each row of the tracker dataframe. I am using comm_or_diss because it already contains everything I need three
upload_file_curve_list = list()
curve_name = set()
for row in range(comm_or_diss.shape[0]):
    if comm_or_diss[row][4] not in curve_name:
        one_curve = query_to_use(comm_or_diss[row][0],comm_or_diss[row][1], comm_or_diss[row][2],comm_or_diss[row][3])
        final_one_curve = one_curve.loc[wk_list(comm_or_diss[row][2], comm_or_diss[row][3])]
        final_one_curve["CURVE_ID"] = comm_or_diss[row][4]
        final_one_curve["SELLING_CYCLE"] = selling_cycle_mapper(comm_or_diss[row][5])
        final_one_curve["PERIOD_NUM"] = np.arange(1,len(final_one_curve)+1)
        final_one_curve["PERIOD_RATIO"] = final_one_curve["Avg2019-2021ratio"]
        final_curve = final_one_curve[["CURVE_ID","SELLING_CYCLE","PERIOD_NUM","PERIOD_RATIO"]]
        upload_file_curve_list.append(final_curve)
    curve_name.add(comm_or_diss[row][4])
#combine a list of curves together
upload_file_curve_list_pandas = pd.concat(upload_file_curve_list)
upload_file_curve_list_pandas.to_excel(os.path.expanduser("~\\OneDrive - Canadian Tire\\Desktop\\UDT_SELLINGCURVE_" + os.getlogin() + "_"
                   + date.today().strftime("%b") + date.today().strftime("%d") + ".xlsx"), index = False)
#change the date format
HIST_MOD_Template_TRACKER['MODEL_START_DATE'] = HIST_MOD_Template_TRACKER['MODEL_START_DATE'].dt.strftime('%Y/%m/%d')
HIST_MOD_Template_TRACKER.loc[:,"STYLE":"RUN_FLAG"].to_excel(os.path.expanduser("~\\OneDrive - Canadian Tire"
                                                        "\\Desktop\\UDT_HISTMODESETUP_F22_TEMPLATE_" + os.getlogin() + ".xlsx"), index = False)

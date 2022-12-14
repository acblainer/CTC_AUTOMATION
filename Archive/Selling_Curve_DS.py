import os
import subprocess
import sys
import pkg_resources
from collections import defaultdict
from datetime import date
import random
import tkinter as tk
from tkinter import filedialog

#July 26 2022: I will make a function out of all the following procedure
#check if you have all the package, if not install it
def selling_curve_prep():
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
    if not os.path.exists('../instantclient_21_6'):
        wget.download(download_link, out = '../')
        with ZipFile('../instantclient-basic-windows.x64-21.6.0.0.0dbru.zip', 'r') as zipObj:
            # Extract all the contents of zip file in current directory
            zipObj.extractall()

## Load the XX XX_HIST MOD Template
# The xlsx is stored at K:\Logistics\_DFP Reports - Menswear\2022\F22 Setup
# NOTE: sometimes the template has missing value for K column for example, you want to double chcck if the plate excel sheet is ok before you run the following code.
def read_query_output():
    import cx_Oracle
    from sqlalchemy.engine import create_engine
    import pandas as pd
    import numpy as np
    from datetime import date

    #prepare 
    lib_dir = '../instantclient_21_6'
    name_string = ""
    for name in os.listdir(lib_dir):
        name_string = name_string + name + " "
    #you just need to run the following snipet for once.
    try:
        cx_Oracle.init_oracle_client(lib_dir=lib_dir)
    except:
        pass
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    #Read the excel file and get a list of sheets. Then chose and load the sheets.
    HIST_MOD_Template = pd.ExcelFile(file_path)
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

    #Connect to the Oracle database using SQLAIchemy (recommended)
    DIALECT = 'oracle'
    SQL_DRIVER = 'cx_oracle'
    USERNAME = "ypfu" #enter your username
    PASSWORD = "cT4K3tgta8NnS!QK" #enter your password
    HOST = "p9cpwpjdadb01" #enter the oracle db host url
    PORT = 25959 # enter the oracle port number
    SERVICE = "FR01PR" # enter the oracle db service name
    ENGINE_PATH_WIN_AUTH = DIALECT + '+' + SQL_DRIVER + '://' + USERNAME + ':' + PASSWORD +'@' + HOST + ':' + str(PORT) + '/?service_name=' + SERVICE

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

    #define query to pull data:
    selling_curve_query_comm = '''
        Select c.CALCWY WEEK,c.CALCYR, sum(nvl(h.qty,0)) as actual_qty
        From scpomgr.cds_histview h
        Inner Join DFREPORTING.CALDTLEE c on h.startdate = c.CALCUR
        --Where h.dmdunit in (Select distinct dmdunit from scpomgr.dmdunit where u_style in ('S_2DIADHAS-419') and dmdunit like 'C_%')
        --Where h.dmdunit in (Select distinct dmdunit from scpomgr.dmdunit where u_dissection = :diss and dmdunit like 'C_%')
        Where h.dmdunit in (Select distinct dmdunit from scpomgr.dmdunit where u_commodity = :comm and dmdunit like 'C_%')
        And h.loc = 'ALL'
        And h.event = 'TOTAL'
        Group By c.CALCWY,c.CALCYR
        Order By c.CALCWY, c.CALCYR
    '''
    selling_curve_query_diss = '''
         Select c.CALCWY WEEK,c.CALCYR, sum(nvl(h.qty,0)) as actual_qty
        From scpomgr.cds_histview h
        Inner Join DFREPORTING.CALDTLEE c on h.startdate = c.CALCUR
        --Where h.dmdunit in (Select distinct dmdunit from scpomgr.dmdunit where u_style in ('S_2DIADHAS-419') and dmdunit like 'C_%')
        Where h.dmdunit in (Select distinct dmdunit from scpomgr.dmdunit where u_dissection = :diss and dmdunit like 'C_%')
        --Where h.dmdunit in (Select distinct dmdunit from scpomgr.dmdunit where u_commodity = :comm and dmdunit like 'C_%')
        And h.loc = 'ALL'
        And h.event = 'TOTAL'
        Group By c.CALCWY,c.CALCYR
        Order By c.CALCWY, c.CALCYR
    '''

    #define a function to decide which query to use for the input weeks
    def query_to_use(holder, indicator, wk1, wk2):
        engine = create_engine(ENGINE_PATH_WIN_AUTH)
        if indicator:
            returned_result = pd.read_sql_query(selling_curve_query_diss, engine, params = [holder])
        else:
            returned_result = pd.read_sql_query(selling_curve_query_comm, engine, params = [holder])
        #dont forget to close the engine to release the resource you have used on the server.
        engine.dispose()
        #remove the last row because we are only considering 52 weeks
        # returned_result = returned_result[:-1]
        # returned_result["Avg2019-2021"] = returned_result[["year_2019", "year_2021"]].mean(axis = 1).apply(np.ceil).astype(int)
        returned_result.loc[returned_result["week"].isin(np.setdiff1d(np.arange(1,53),wk_list(wk1, wk2))), "Avg2019-2021"] = 0
        # returned_result["Avg2019-2021ratio"] = (returned_result["Avg2019-2021"]/returned_result["Avg2019-2021"].sum())
        # returned_result.index = np.arange(1,len(returned_result) + 1)
        return returned_result

    #scan through each row of the tracker dataframe. I am using comm_or_diss because it already contains everything I need three
    upload_file_curve_list = list()
    curve_name = set()
    for row in range(comm_or_diss.shape[0]):
        if comm_or_diss[row][4] not in curve_name:
            one_curve = query_to_use(comm_or_diss[row][0],comm_or_diss[row][1], comm_or_diss[row][2],comm_or_diss[row][3])
            one_curve['curve_name'] = comm_or_diss[row][4]
            upload_file_curve_list.append(one_curve)
        curve_name.add(comm_or_diss[row][4])
    #combine a list of curves together
    upload_file_curve_list_pandas = pd.concat(upload_file_curve_list)
    upload_file_curve_list_pandas.to_excel(os.path.expanduser("~\\OneDrive - Canadian Tire\\Desktop\\UDT_SELLINGCURVE_" + os.getlogin() + "_"
                    + date.today().strftime("%b") + date.today().strftime("%d") + ".xlsx"), index = False)

if __name__ == "__main__":
    selling_curve_prep()
    read_query_output()

import os
import subprocess
import sys
import pkg_resources
from collections import defaultdict
from datetime import date
import random

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
    if not os.path.exists('./instantclient_21_6'):
        wget.download(download_link)
        with ZipFile('./instantclient-basic-windows.x64-21.6.0.0.0dbru.zip', 'r') as zipObj:
            # Extract all the contents of zip file in current directory
            zipObj.extractall()

## Load the XX XX_HIST MOD Template
# The xlsx is stored at K:\Logistics\_DFP Reports - Menswear\2022\F22 Setup
# NOTE: sometimes the template has missing value for K column for example, you want to double chcck if the plate excel sheet is ok before you run the following code.
def read_query_output(file_path, **kargs):
    import cx_Oracle
    from sqlalchemy.engine import create_engine
    import pandas as pd
    import numpy as np
    from datetime import date

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
    DIALECT = kargs['DIALECT'] #'oracle'
    SQL_DRIVER = kargs['SQL_DRIVER']#cx_oracle'
    USERNAME = kargs['USERNAME']#"ypfu" #enter your username
    PASSWORD = kargs['PASSWORD']#"cT4K3tgta8NnS!QK" #enter your password
    HOST = kargs['HOST']#"p9cpwpjdadb01" #enter the oracle db host url
    PORT = kargs['PORT']#25959 # enter the oracle port number
    SERVICE = kargs['SERVICE']#"FR01PR" # enter the oracle db service name
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
        engine = create_engine(ENGINE_PATH_WIN_AUTH)
        if indicator:
            returned_result = pd.read_sql_query(selling_curve_query_diss, engine, params = [holder])
        else:
            returned_result = pd.read_sql_query(selling_curve_query_comm, engine, params = [holder])
        #dont forget to close the engine to release the resource you have used on the server.
        engine.dispose()
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
            #it will make a difference when I using AS-All Season for Selling Cycle in the selling curve builder. 
            # I guess it is taking all year ratio as the final result, regardless of the Instore week and Endweek
            if (comm_or_diss[row][5] == 'All Season'):
                final_one_curve = one_curve
            else:
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


#prepare the Consolidation
#check if you have all the package, if not install it
def consolidation_prep():
    def install(package):
        if (package not in {pkg.key for pkg in pkg_resources.working_set}):
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
    for pkg in ['pandas', 'numpy']:
        install(pkg)
def consolidation_output(file_path):
    import pandas as pd
    import numpy as np
    class Store:
        #The information for all stores combined
        all_store = None
        __all_store_flag = False
        #record what is the total capacity each store can hold
        store_capacity = None
        #one particular store can only accpet 250 total transferred units
        store_total_addition = defaultdict(lambda:0)
        
        @classmethod #I make this private because I dont want it exposed outside of this class
        def __find_match(cls, all_store_sku, sending_store_sku):
            match_dict = {}
            for k, v in all_store_sku[all_store_sku['PA for the year'] > 0].groupby('PA_Store')['Sku']:
                match_dict[k] = len(list(set(v)&set(sending_store_sku.loc[sending_store_sku.OH > 0,'Sku'])))
            return match_dict
        
        #initiate the store with a pandas dataframe
        def __init__(self, one_store, all_store):
            #I will only accpet all_store variable for the fitst time instance creation
            if not Store.__all_store_flag:
                Store.all_store = all_store.copy()
                Store.__all_store_flag = True
                Store.store_capacity = Store.all_store.groupby('PA_Store')['PA for the year'].sum()
            #accept the shipping store
            self.one_store = one_store.copy()
            #save one original copy of the store
            self.one_store_copy = one_store.copy()
            #only look at all the stores that have the same region of this sending store
            self.all_store_same_region = Store.all_store.loc[Store.all_store.Region == self.one_store.loc[0, 'Region']]
            #keep a record how many stores (B) has been used for transferring this store (A)
            self.total_out_id = 0
            #keep track which store has been used to receive the transfers
            self.receiving_store = {}
            #what is the matching number for this sending store with every single other stores there
            self.matching_series = pd.Series(Store.__find_match(self.all_store_same_region, self.one_store),
                                            name = "matched_series_same_region")
            #the rank of each store based on the total capacity of a store and weighted with the percentage of matching number
            #this variable has excluded the sending store ID (we dont need the rank of itself)
            #NOTE I only calculate the stores that have the same Region as sending store
            #what if there is no match at all, and what if some of them are matching 0, some of them have partial match
            self.store_rank = (((Store.store_capacity[list(set(self.all_store_same_region.PA_Store))]) * \
                                ((self.matching_series/self.matching_series.sum()).where(~(self.matching_series/self.matching_series.sum()).isin([0,np.nan,-np.inf, np.inf]),0.0000001))).sort_values(ascending=False))

        #when a list of skus from store (A) are transferred to another store (B):
        #(1) update the SKU (PA for the year in store B) with the corresponding OH from A store
        #(2) change the OH for that SKU in this store (A) to 0
        #store_id is the receiving store B ID
        #the client of this class needs to make sure the sku_list is unique
        def transfer(self, store_id, sku_list):
            #get a subset of info that only contains sku_list for both one_store and all_store
                    #then merge them together, create a new column that returns the minimum of OH (one_store) and 'PA for the year' (all_store) for each SKU
            common_sending_receiving = pd.merge(self.one_store.loc[self.one_store.Sku.isin(sku_list)], Store.all_store.loc[(Store.all_store.PA_Store == store_id) & Store.all_store.Sku.isin(sku_list)],
                                                on = 'Sku', how = 'inner')
            common_sending_receiving['min_transfer'] = common_sending_receiving[['OH', 'PA for the year']].min(axis = 1)
            sku_list_map = dict(zip(common_sending_receiving.Sku, common_sending_receiving.min_transfer))
            #if the total SKUs transferred is less than 4 units or less, I do not want that to transfer
            if common_sending_receiving['min_transfer'].sum() <= 4:
                sku_list_map = {}
            #use this column to serve as a mapper to update all_store (first), all_store_same_region (then), store_capacity (then), store_total_addition (then), and finally one_store itself
            tem_all_store = Store.all_store.loc[Store.all_store.PA_Store == store_id, 'Sku'].map(sku_list_map).fillna(0).astype(int)
            Store.all_store.loc[Store.all_store.PA_Store == store_id, 'PA for the year'] -= tem_all_store

            #then update the all_store_same_region as well based on the new all_store
            self.all_store_same_region = Store.all_store.loc[Store.all_store.Region == self.one_store.loc[0, 'Region']]
            
            #update the total capacity each receiving store can hold once it receives some items
            Store.store_capacity = Store.all_store.groupby('PA_Store')['PA for the year'].sum()
            #one particular store can only accpet 250 total transferred units
            Store.store_total_addition[store_id] += common_sending_receiving['min_transfer'].sum()
        
            #update one_store A
            tem_one_store = self.one_store.loc[:,'Sku'].map(sku_list_map).fillna(0).astype(int)
            self.one_store.loc[self.one_store.Sku.isin(sku_list),'OH'] -= tem_one_store
            
            #keep a record which score B has received what SKU list
            #when updating receiving_store, I will use tuple (one_store_id, receiving_store_id, sku_id) as the dict key: how_many_is_transferred
            for key, value in sku_list_map.items():
                self.receiving_store[(self.one_store.loc[0,"Sending_Store"],store_id, key)] = value
            
            #keep a record how many stores (B) has been used for transferring this store (A)
            self.total_out_id = len(set([key[1] for key in self.receiving_store.keys()]))
            
            #what is the matching number for this sending store with every single other stores there
            self.matching_series = pd.Series(Store.__find_match(self.all_store_same_region, self.one_store),
                                            name = "matched_series_same_region")
            
            #this variable excludes the sending store ID because Store.all_store has dropped the senidng store ID (we dont need the rank of itself)
            #NOTE I only calculate the stores that have the same Region as sending store
            #what if there is no match at all, and what if some of them are matching 0, some of them have partial match
            #if there is no match (0) for individual store, or there is no match at all between sending stores and receiving stores (NaN)
            #I will give them a really small weight, 0.0000001
            self.store_rank = (((Store.store_capacity[list(set(self.all_store_same_region.PA_Store))]) * \
                                ((self.matching_series/self.matching_series.sum()).where(~(self.matching_series/self.matching_series.sum()).isin([0,np.nan,-np.inf, np.inf]),0.0000001))).sort_values(ascending=False))
            
    #Transfer ONE sending store A to a list of receiving stores B
    def consolidation(sending_store, all_store):
        #(1): initiate the store object
        sending_store = Store(sending_store, all_store)

        #(2): rank the stores based on the capacity of each store but add weight to each store 
        #based on the percentage of each matching nunber
        #NOTE: store_rank only looks at the stores that have the same region as sending store
        
        #(3): going through the receiving store in this order and 
        #if one store has not received 250 units in total before, starts transferring.
        total_receiving_number = 750
        #every time a store is used to receive some SKUs, I will remove it from the store_rank_list
        removed_store_id = []
        #current_store_rank_list is what remained after removing the used store
        current_store_rank_list = list(sending_store.store_rank.index)
        #there are 2 conditions for me to keep transferrring. one is there is still some SKU on hands, another one is there are still some remaining store to receive 
        while (sending_store.one_store['OH'].sum() > 0 and  len(current_store_rank_list) > 0):
            #when the current rank store has not reached 250, we will transfer some over
            if(sending_store.store_total_addition[current_store_rank_list[0]] < total_receiving_number):
            #find the common SKUs (the OH or PA for the year must >0, otherwise there is no point transferring) 
            #between sending store and current receiving store, and then start transferring
                sending_sku_set = set(sending_store.one_store.loc[sending_store.one_store['OH'] > 0, 'Sku']) \
                & set(sending_store.all_store_same_region.loc[(sending_store.all_store_same_region.PA_Store == current_store_rank_list[0]) & \
                                                (sending_store.all_store_same_region['PA for the year'] > 0), 'Sku'])
                #start transferring
                sending_store.transfer(current_store_rank_list[0],list(sending_sku_set))
                #after transferring, I will keep a record of which store I want to remove
                removed_store_id.append(current_store_rank_list.pop(0))
                #I will remove this store from the store_rank for the next iteration, but I still keep the order of the new ranking
                current_store_rank_list = [elem for elem in list(sending_store.store_rank.index) if elem not in removed_store_id]
                
            else:
                #although there is no transfer happened, I will still need to remove this store from the ranking list
                #because it tells me that this store has reached 250, we cannot use this any way
                removed_store_id.append(current_store_rank_list.pop(0))
                current_store_rank_list = [elem for elem in list(sending_store.store_rank.index) if elem not in removed_store_id]
        #at last, we return the sending store info regardless if all SKUs are transferred out or not
        return sending_store

    # create the function to output the transferred info
    #consolidated_stores is a list here
    #also consider if there is any SKUs not transferred??????
    def output_stores(consolidated_stores_list):
        store_list = []
        for store in consolidated_stores_list:
            multi_index = pd.MultiIndex.from_tuples(store.receiving_store.keys(), 
                                                    names=["Sending_Store", "Receiving_Store", 'Sku'])
            store_transferred = pd.DataFrame(store.receiving_store.values(), 
                                            index = multi_index, columns = ['Qty']).reset_index()
            
            store_list.append(store_transferred)
        return pd.concat(store_list)

    #read in the excel sheet
    #NOTE: when you try to read excel in python, you should close any excel file you want to load first
    Sending_Info = pd.read_excel(file_path, sheet_name = "Sending Info").dropna(how='any', subset = ['Sending_Store', 'Region'])
    PA_Info = pd.read_excel(file_path, sheet_name = "PA Info").dropna(how='any', subset = ['PA_Store', 'Region'])
    Region_info = pd.read_excel(file_path, sheet_name = "Region", usecols = "A,F,K, O",skiprows = range(0,4), header = 0)
    Region_info_hub = Region_info.loc[Region_info['Hub Designation'].str.contains('High|Medium', na=False,regex = True)].reset_index(drop = True)
    Region_info_province_dict = Region_info_hub.groupby('Province')['STR'].apply(lambda g: g.values.tolist()).to_dict()
    #Create a dictionary for each store to have a corresponing province
    Region_info_store_province_dict = Region_info.loc[:,["STR", "Province"]].set_index("STR").to_dict()['Province']
    #group same Sku together for both info sheet
    Sending_Info_group = Sending_Info.groupby(['Sending_Store','Sku','Region'],as_index = False)['OH'].sum()
    #remove the sending store id in the PA_Info sheet
    #PA_Info = PA_Info.loc[~PA_Info.PA_Store.isin(set(Sending_Info['Sending_Store']))]
    PA_Info_group = PA_Info.groupby(['PA_Store','Sku','Region'],as_index = False)['PA for the year'].sum()

    #do the actual consolidation store by store
    sending_store_list = []
    #deal with the reamining skus
    remaining_sku_list = []
    for key, group in Sending_Info_group.groupby('Sending_Store'):
        sending_storeA = consolidation(group.reset_index(drop = True), PA_Info_group)
        sending_store_list.append(sending_storeA)
        if(len(sending_storeA.one_store.query('OH>0')) >0):
            #Store 181 and 193 should go to hub store 46, 59, 301 those are the closest
            if sending_storeA.one_store.loc[0,'Sending_Store'] in [181, 193]:
                hub_store_same_province = [46, 59, 301]
            #Newfoundland can go to any Nova Scotia or New Brunswick hub
            elif Region_info_store_province_dict[sending_storeA.one_store.loc[0,'Sending_Store']] == 'Newfoundland':
                hub_store_same_province = list(np.concatenate([value for key, value in Region_info_province_dict.items() if key in ['Nova Scotia', 'New Brunswick']]).flat)
            else:
                hub_store_same_province = Region_info_province_dict[Region_info_store_province_dict[sending_storeA.one_store.loc[0,'Sending_Store']]]
            #I shuffle the hub store order in this particular region so that one hub store does not end up receiving all the SKUs
            random.shuffle(hub_store_same_province)
            hub_store_same_province_output = sending_storeA.one_store.loc[sending_storeA.one_store['OH'] > 0].reset_index(drop = True)
            hub_store_same_province_output.rename(columns={'OH': 'Qty'}, inplace = True)
            hub_store_same_province_output.loc[:,'Receiving_Store'] = hub_store_same_province[0]
            remaining_sku_list.append(hub_store_same_province_output.loc[:,['Sending_Store', 'Receiving_Store', 'Sku', 'Qty']])

    #generate the result
    output_stores(sending_store_list).to_excel(os.path.expanduser("~\\OneDrive - Canadian Tire\\Desktop\\sending_store_list_12mth_750_with4_allstore" + os.getlogin() + "_"
                    + date.today().strftime("%b") + date.today().strftime("%d") + ".xlsx"), index = False)
    pd.concat(remaining_sku_list).to_excel(os.path.expanduser("~\\OneDrive - Canadian Tire\\Desktop\\remaining_sku_list_12mth_750_with4_allstore" + os.getlogin() + "_"
                    + date.today().strftime("%b") + date.today().strftime("%d") + ".xlsx"), index = False)
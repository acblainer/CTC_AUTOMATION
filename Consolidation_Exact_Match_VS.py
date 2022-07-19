#load the necessary library
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
from collections import defaultdict
import os
from datetime import date
import random
#we need to remove all the sending store id in the receiving store B list

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
    total_receiving_number = 500
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
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
Sending_Info = pd.read_excel(file_path, sheet_name = "Sending Info").dropna(how='all', subset = ['Sending_Store', 'Region'])
PA_Info = pd.read_excel(file_path, sheet_name = "PA Info").dropna(how='all', subset = ['PA_Store', 'Region'])
Region_info = pd.read_excel(file_path, sheet_name = "Region", usecols = "A,F,K",skiprows = range(0,4), header = 0)
Region_info_hub = Region_info.loc[Region_info['Hub Designation'].str.contains('High|Medium', na=False,regex = True)].reset_index(drop = True)
Region_info_hub_dict = Region_info_hub.groupby('Region')['STR'].apply(lambda g: g.values.tolist()).to_dict()
#group same Sku together for both info sheet
Sending_Info_group = Sending_Info.groupby(['Sending_Store','Sku','Region'],as_index = False)['OH'].sum()
#remove the sending store id in the PA_Info sheet
PA_Info = PA_Info.loc[~PA_Info.PA_Store.isin(set(Sending_Info['Sending_Store']))]
PA_Info_group = PA_Info.groupby(['PA_Store','Sku','Region'],as_index = False)['PA for the year'].sum()

#do the actual consolidation store by store
sending_store_list = []
#deal with the reamining skus
remaining_sku_list = []
for key, group in Sending_Info_group.groupby('Sending_Store'):
    sending_storeA = consolidation(group.reset_index(drop = True), PA_Info_group)
    sending_store_list.append(sending_storeA)
    if(len(sending_storeA.one_store.query('OH>0')) >0):
        hub_store_same_region = Region_info_hub_dict[sending_storeA.one_store.loc[0,'Region']]
        #I shuffle the hub store order in this particular region so that one hub store does not end up receiving all the SKUs
        random.shuffle(hub_store_same_region)
        hub_store_same_region_output = sending_storeA.one_store.loc[sending_storeA.one_store['OH'] > 0].reset_index(drop = True)
        hub_store_same_region_output.rename(columns={'OH': 'Qty'}, inplace = True)
        hub_store_same_region_output.loc[:,'Receiving_Store'] = hub_store_same_region[0]
        remaining_sku_list.append(hub_store_same_region_output.loc[:,['Sending_Store', 'Receiving_Store', 'Sku', 'Qty']])

#generate the result
output_stores(sending_store_list).to_excel(os.path.expanduser("~\\OneDrive - Canadian Tire\\Desktop\\sending_store_list_6mth_500" + os.getlogin() + "_"
                   + date.today().strftime("%b") + date.today().strftime("%d") + ".xlsx"), index = False)
pd.concat(remaining_sku_list).to_excel(os.path.expanduser("~\\OneDrive - Canadian Tire\\Desktop\\remaining_sku_list_6mth_500" + os.getlogin() + "_"
                   + date.today().strftime("%b") + date.today().strftime("%d") + ".xlsx"), index = False)
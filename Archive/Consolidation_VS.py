#load the necessary library
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
from collections import defaultdict
import os
from datetime import date

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
        self.receiving_store = defaultdict(list)
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
        
        #update store B in the all_store info sheet
        sku_list_map = self.one_store[self.one_store.Sku.isin(sku_list)].set_index('Sku')['OH'].to_dict()
        tem = Store.all_store.loc[Store.all_store.PA_Store == store_id, 'Sku'].map(sku_list_map).fillna(0)
        Store.all_store.loc[Store.all_store.PA_Store == store_id,'PA for the year'] -= tem

        #then update the all_store_same_region as well based on the new all_store
        self.all_store_same_region = Store.all_store.loc[Store.all_store.Region == self.one_store.loc[0, 'Region']]
        
        #update the total capacity each receiving store can hold once it receives some items ??
        Store.store_capacity = Store.all_store.groupby('PA_Store')['PA for the year'].sum()
        #one particular store can only accpet 250 total transferred units
        Store.store_total_addition[store_id] += self.one_store[self.one_store.Sku.isin(sku_list)].set_index('Sku')['OH'].sum()
           
        #update store A (this must be done after updating B)
        self.one_store.loc[self.one_store.Sku.isin(sku_list),'OH'] = 0
        
        #keep a record which score B has received what SKU list
        self.receiving_store[store_id].extend(sku_list)
        #keep a record how many stores (B) has been used for transferring this store (A)
        #one store A can only be split into 2 stores B
        self.total_out_id = len(self.receiving_store.keys())
        
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

#Transfer one sending store A to a list of receiving stores B
def consolidation(sending_store, all_store):
    #(1): initiate the store object
    sending_store = Store(sending_store, all_store)

    #(2): rank the stores based on the capacity of each store but add weight to each store 
    #based on the percentage of each matching nunber
    #NOTE: store_rank only looks at the stores that have the same region as sending store
    
    #(3): going through the receiving store in this order and 
    #if one store has not received 250 units in total before, starts transferring.
    total_receiving_number = 250
    #filter out all the stores that has the different region as the sending store A
    #I dont explicitly do this because store_rank returns only the stores that have the same region as sending store
    for index, _ in sending_store.store_rank.items():
        if sending_store.store_total_addition[index] < total_receiving_number:
            #(3.1): when the receiving store can receive more than 80% of the sku, transfer all of the skus in store A
            if sending_store.store_capacity[index] >= (sending_store.one_store.OH.sum())*0.8:
                sending_store.transfer(index,sending_store.one_store.Sku.tolist())
                return sending_store
            #(3.2):if less than 80%,transfer as much as the receiving store B can receive,
            #then I will recalcualte the rank of the store after transfer some portion of SKUs out
            #and all the remaining MUST be transferred completely to next available highest receiving store B,
            #in other words, one sending store can only be split into 2 stores maximum
            else:
                sending_sku_set = set(sending_store.one_store.loc[:, 'Sku']) \
                                        & set(sending_store.all_store_same_region.loc[(sending_store.all_store_same_region.PA_Store == index) & \
                                             (sending_store.all_store_same_region['PA for the year'] > 0), 'Sku'])
                sending_store.transfer(index,list(sending_sku_set))
                for index, _ in sending_store.store_rank.items():
                    if sending_store.store_total_addition[index] < total_receiving_number:
                        #transfer the remaining SKUs to this receiving store B
                        remaining_sku_set = set(sending_store.one_store.loc[:, 'Sku']) - sending_sku_set                     
                        sending_store.transfer(index,remaining_sku_set)
                        return sending_store
                    else:
                        continue
        else:
            continue

# create the function to output the transferred info
#consolidated_stores is a list here
def output_stores(consolidated_stores):
    store_list = []
    for store in consolidated_stores:
        tem_store_copy = store.one_store_copy.copy()
        tem_store_output = store.one_store_copy.loc[:,['Sending_Store', 'Sku', 'OH']]
        for store_id,sku_list in store.receiving_store.items():
            tem_store_copy['Sku'] = np.where(tem_store_copy['Sku'].isin(sku_list), store_id, tem_store_copy['Sku'])
        tem_store_output['Receiving_Store'] = tem_store_copy['Sku']
        tem_store_output['Qty'] = tem_store_output['OH']
        tem_store_output['Style'] = np.nan
        tem_store_output['Choice'] = np.nan
        tem_store_output['Cost'] = np.nan
        tem_store_output['Retail'] = np.nan
        final_store_output = tem_store_output[['Sending_Store', 'Receiving_Store', 'Style', 'Choice', 'Sku', 'Qty', 'Cost', 'Retail']]
        store_list.append(final_store_output)
    return pd.concat(store_list)


#read in the excel sheet
#NOTE: when you try to read excel in python, you should close any excel file you want to load first
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
Sending_Info = pd.read_excel(file_path, sheet_name = "Sending Info").dropna(how='all', subset = ['Sending_Store', 'Region'])
PA_Info = pd.read_excel(file_path, sheet_name = "PA Info").dropna(how='all', subset = ['PA_Store', 'Region'])
#group same Sku together for both info sheet
Sending_Info_group = Sending_Info.groupby(['Sending_Store','Sku','Region'],as_index = False)['OH'].sum()
#remove the sending store id in the PA_Info sheet
PA_Info = PA_Info.loc[~PA_Info.PA_Store.isin(set(Sending_Info['Sending_Store']))]
PA_Info_group = PA_Info.groupby(['PA_Store','Sku','Region'],as_index = False)['PA for the year'].sum()

#do the actual consolidation store by store
sending_store_list = []
for key, group in Sending_Info_group.groupby('Sending_Store'):
    sending_storeA = consolidation(group.reset_index(), PA_Info_group)
    sending_store_list.append(sending_storeA)
#generate the result
output_stores(sending_store_list).to_excel(os.path.expanduser("~\\OneDrive - Canadian Tire\\Desktop\\test" + os.getlogin() + "_"
                   + date.today().strftime("%b") + date.today().strftime("%d") + ".xlsx"), index = False)
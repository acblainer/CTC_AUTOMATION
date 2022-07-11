#load the necessary library
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
from collections import defaultdict

#read in the excel sheet
#NOTE: when you try to read excel in python, you should close any excel file you want to load first
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
Sending_Info = pd.read_excel(file_path, sheet_name = "Sending Info")
PA_Info = pd.read_excel(file_path, sheet_name = "PA Info")
#group same Sku together for both info sheet
Sending_Info_group = Sending_Info.groupby(['Sending_Store', 'Comm', 'Style','Sku','Region'],as_index = False)['OH'].sum()
PA_Info_group = PA_Info.groupby(['PA_Store', 'Comm', 'Style','Sku','Region'],as_index = False)['PA for the year'].sum()

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
        #keep a record how many stores (B) has been used for transferring this store (A)
        self.total_out_id = 0
        #keep track which store has been used to receive the transfers
        self.receiving_store = defaultdict(list)
        #what is the matching number for this sending store with every single other stores there
        self.matching_series = pd.Series(Store.__find_match(Store.all_store, self.one_store),
                                        name = "matched_series")
        #the rank of each store based on the total capacity of a store and weighted with the percentage of matching number
        #this variable excludes the sending store ID (we dont need the rank of itself)
        self.store_rank = (((Store.store_capacity) * \
                            (self.matching_series/self.matching_series.sum())).sort_values(ascending=False)).drop(self.one_store.iloc[0,0])
    
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
        #add new value to SKU column when that SKU is not matched? (no need to do this)
        #Lets delete this sending store id in the all_store sheet so that the coming sending store wont be transferred to the old sneding store in the master sheet
        Store.all_store.drop(Store.all_store[Store.all_store.PA_Store == self.one_store.iloc[0,0]].index, inplace=True)
        
        #update the total capacity each receiving store can hold once it receives some items
        Store.store_capacity = Store.all_store.groupby('PA_Store')['PA for the year'].sum()
        Store.store_capacity[store_id] -= tem.sum()
        #one particular store can only accpet 250 total transferred units
        Store.store_total_addition[store_id] += tem.sum()
           
        #update store A (this must be done after updating B)
        self.one_store.loc[self.one_store.Sku.isin(sku_list),'OH'] = 0
        
        #keep a record which score B has received what SKU list
        self.receiving_store[store_id].extend(sku_list)
        #keep a record how many stores (B) has been used for transferring this store (A)
        #one store A can only be split into 2 stores B
        self.total_out_id = len(self.receiving_store.keys())
        
        #what is the matching number for this sending store with every single other stores there
        self.matching_series = pd.Series(Store.__find_match(Store.all_store, self.one_store),
                                        name = "matched_series")
        
        #this variable excludes the sending store ID because Store.all_store has dropped the senidng store ID (we dont need the rank of itself)
        self.store_rank = (((Store.store_capacity) * \
                            (self.matching_series/self.matching_series.sum())).sort_values(ascending=False))


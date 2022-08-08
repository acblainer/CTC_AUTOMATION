import pandas as pd
import numpy as np
import dask.dataframe as dd

col_list_reduced = ['SALES_DATE', 'Transaction Number', 'Norm Sub Channel',
                    'SKU_NUM', 'SKU_NAME', 'STYLE_NO', 'ITEM_QTY', 'SHIPPED_TO_POSTALCODE', 'SHIPPED_TO_CITY', 'SHIPPED_TO_PROVINCE']
dtypes = {"Transaction Number":"str", 'Norm Sub Channel':"str",
          "SKU_NUM":"str", 'SKU_NAME':"str", "STYLE_NO":"str", "ITEM_QTY":"int",
         'SHIPPED_TO_POSTALCODE':"str", 'SHIPPED_TO_CITY':"str", 'SHIPPED_TO_PROVINCE':"str"}
#read in the csv
OCTOBER_2021_reduced_web_ship_pre = pd.read_csv(r"S:/Olivia/Txn Data Extracts from Akshay/OCTOBER_2021_NEW_LOGIC_USING_UNITS_REFRESHED_MDR_SORTED.csv",
                          usecols = col_list_reduced, engine = 'c',dtype=dtypes, parse_dates = ["SALES_DATE"], encoding='windows-1252', chunksize = 10000000)
OCTOBER_2021_reduced_web_ship_final = pd.concat([chunk.query("`Norm Sub Channel` == 'GROSS WEB SALES'") for chunk in OCTOBER_2021_reduced_web_ship_pre], ignore_index = True)

NOVEMBER_2021_reduced_web_ship_pre = pd.read_csv(r"S:/Olivia/Txn Data Extracts from Akshay/NOVEMBER_2021_NEW_LOGIC_USING_UNITS_REFRESHED_MDR_SORTED.csv",
                          usecols = col_list_reduced, engine = 'c',dtype=dtypes, parse_dates = ["SALES_DATE"],encoding='windows-1252',chunksize = 10000000)
NOVEMBER_2021_reduced_web_ship_final = pd.concat([chunk.query("`Norm Sub Channel` == 'GROSS WEB SALES'") for chunk in NOVEMBER_2021_reduced_web_ship_pre], ignore_index = True)

DECEMBER_2021_reduced_web_ship_pre = pd.read_csv(r"S:/Olivia/Txn Data Extracts from Akshay/DECEMBER_2021_NEW_LOGIC_USING_UNITS_REFRESHED_MDR_SORTED.csv",
                          usecols = col_list_reduced, engine = 'c',dtype=dtypes, parse_dates = ["SALES_DATE"],encoding='windows-1252',chunksize = 10000000)
DECEMBER_2021_reduced_web_ship_final = pd.concat([chunk.query("`Norm Sub Channel` == 'GROSS WEB SALES'") for chunk in DECEMBER_2021_reduced_web_ship_pre], ignore_index = True)

Q3_2021_reduced_web_ship_pre = pd.read_csv(r"S:/Olivia/Txn Data Extracts from Akshay/Q3_2021_NEW_LOGIC_USING_UNITS_REFRESHED_MDR_SORTED.csv",
                          usecols = col_list_reduced, engine = 'c',dtype=dtypes, parse_dates = ["SALES_DATE"],encoding='windows-1252',chunksize = 10000000)
Q3_2021_reduced_web_ship_final = pd.concat([chunk.query("`Norm Sub Channel` == 'GROSS WEB SALES'") for chunk in Q3_2021_reduced_web_ship_pre], ignore_index = True)

Fall_2021_reduced_web_ship_final = pd.concat([Q3_2021_reduced_web_ship_final, OCTOBER_2021_reduced_web_ship_final, NOVEMBER_2021_reduced_web_ship_final, DECEMBER_2021_reduced_web_ship_final])
Fall_2021_reduced_web_ship_final.to_csv(r"K:/Logistics/Co-op/Yongpeng/2021 Transaction Data from Olivia Q3 and Fall/Fall_2021_reduced_web_ship_final.csv", index = False)
# CTC_AUTOMATION
We are initializing a project for automating massive manual work at CTC.  
A small hack: how to install and import one package using script [linke](https://stackoverflow.com/questions/12332975/installing-python-module-within-code)


## Name a list of tasks you think you want to automate.
**Yong** 
- Selling Curve: I am in the process of creating a bot (90% work done)
- Some comments from Christy.
  - Make the code more flexible to accept user input
  - The selling cycle does not matter in the SQL query, its main purpose is for curve name convention
  - Sometimes different years have more than 52 weeks, you need to consider that as well
  - Looking at the selling curve at the style level can make the prediction better
  - use machine learning to group similar styles together so that the same selling curve can be applied to them
  - only build the selling curve for the styles that has the exact start week.
  - Question 1: what is dmdunit and why it has to start with C_
  - Question 2: implement the season (S1,S2,S3,S4) based on the season we are in right now
  - QUestion 3: do we need to remove duplicates when copy and paste the yellow portion of the Tracker to "2A_CE_UDT_HISTMODESETUP_F22 TEMPLATE FOR YONG"
  - DIscussion with Leszek (july 27 2022): 
    - find the outliers of the selling curve
    - maybe suggest better way to make selling curve, say different stores have their own customized selling curve
- Create the tool for store consolidation (july 7 2022)
  - I have created the Store class to track the information of a store, including its ranking in the whole store list, and making transfer between this store and the whole store list
  - Now I am implementing the logic like when and how I should transfer the store item
  - **It will make a difference when you choose AS_All Season vs the others**
  
  
**Adam**: (not tracked)
- (July 7 2022): Heres my notebook file for the recship updates. I noticed a small part of the code doesnt work on the anaconda I personally downloaded because I originally built it on the company anaconda which has an old version of pandas. It might not work for you but I'm pretty sure its an easy fix, it just has something to do with the fillna code.


## Update for consolidaiton: find exact match for consolidation
After talking with Linda onJuly 14 2022, we have to modify the consolidation rules. The rule from last time (Consolidation_VS) is called fuzzy match.
**Yong July 16 2022**

Now we need to modify the code so that all the skus have a 100% match.

**How to find the best receiving store B for sending store A.**
- NOTE: each sku is unique....in that they are attached to a style number...think SKU = SIZE
- Find all the stores B that have the same region (East/West) as store A
- Remove all the sending stores A ID in the receiving stores B list (you dont have to worry about this because the SKUs in the sending stores has no overlap with other sending stores.
- (1) **Each time** you do transfer, re-rank each store B based on the each total PA and with different weight on them. 
    - The weight is calculated as follows:
    > Step 1: find the how many SKUs (that are > 0 OH or PA for the year) are matched between store A with each store B
    
    > Step 2: calculate the percentage of each matching number (what if there is no match at all, and what if some of them are matching 0, some of them have partial match. if there is no match (0) for individual store, or there is no match at all between sending stores and receiving stores (NaN). I will give them a really small weight, 0.0000001)
    
    > Step 3: apply this percentage (weight) to the each total PA for each store B
    
- (2) Loop through each sending store A and find the highest rank for each of them, and start transfering with the following conditions:
    - Rule #1: one particular store B can only accept 250 units lifetime maximum
    - Rule #2: transfer the matched SKUs from sending store A to Store B, but only transfer as much Store B can accpet for one particular SKU. However, if there are 4 units or less to be transferred to one particular store, STOP. Do not transfer.
    - Rule #3: the remaining SKUs will go through (1), find the best Store B to transfer, then again only  transfer the matched SKUs from sending store A to Store B, but only transfer as much Store B can accpet for one particular SKU.
    - Rule #4: you keep going this process until you finish up all the SKUs, or you reach the end of receiving stores B, whichever comes first.
   
- (3) I would use only high and medium hub stores only.  And region still matters when choosing which hub stores you send remaining transfers (4 units and less +units you could not find a match). An East store can only send to a East hub store



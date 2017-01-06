# -*- coding: utf-8 -*-
"""
Created on Mon Nov 28 15:57:18 2016

@author: dh1023
"""

# For a given data set, find how many weeks there are then find avg by 
# priority and craft and add est hours to a dataframe

# Import libraries
import pandas as pd
#import matplotlib
import numpy as np
#import sys
#import matplotlib
from datetime import datetime, timedelta, date
import datetime as dt
import math

'''
Import the data
'''
# Location of file with previous SR data
# This data can have any length of date range and should include all WO's
# This range of data usually mirrors the time period for which it is seeking
# to predict: i.e. if the next two weeks are Dec 1-14, 2016, then the data
# should cover Dec 1-14, 2015 and potentially some buffer on either side.
# This is the raw data used to predict future SR.
sr_excel = r'C:\Users\dh1023.AD\Desktop\Python\Zone_Manager_Report_SR_Project' \
         'ed_-_DGH_Master.xlsx'

# Create dataframe from previous SR data
sr_df = pd.read_excel(sr_excel)

# Location of file with data of what buildings belong to what zones
# Note that first column should be "Bldg/ Land Entity" which is the building 
# number and second column should be "Zone" 
bldg_excel = r'C:\Users\dh1023.AD\Desktop\Python\FMS61100_-_Active_Build' \
               'ing_or_Land_Entity_Listing.xlsx'

# Create dataframe from building data
bldg_df = pd.read_excel(bldg_excel)





'''
Merge/Join the Zone and Building Information to SR data
'''
# change the name of the bldg number to make merge/join easier
bldg_df.columns = ['WO Building', 'Zone']

# make sure the building numbers are numbers, not text from WebI
sr_df['WO Building'] = pd.to_numeric(sr_df['WO Building'], errors='coerce')
bldg_df['WO Building'] = pd.to_numeric(bldg_df['WO Building'], errors='coerce')

# Merge/join data from Buldings so that the work orders have the correct 
# current zone/ops associated with them
sr_df = sr_df.merge(bldg_df, on='WO Building', how='left')





'''
Find the average per week of work orders and use as our projection
'''
# how many weeks worth of data do we have?
weeks_of_data = (sr_df['Enter Date'].max() - sr_df['Enter Date'].min()).days/7

# How many WO's were entered by priority, craft, & zone over the given period?
# Use ".size()" instead of .count() because size counts null values.
# Note: our historical average becomes our projected work.
previous_wo = sr_df.groupby(['WO Crew','Zone','Craft_v','WO Priority']).agg({
                            'WO Num': np.size, 'Est Hrs WO Calculated_v': 
                                np.sum})

# rename WO Num column to "WOcount" since it no longer is individual WO's
# but a summation of WO's
previous_wo = previous_wo.rename(columns = {'WO Num':'WOcount'})

# what does that average out to be per week?
avg_wo_per_wk = previous_wo / weeks_of_data

# reset index to actually create a dataframe
avg_wo_per_wk = avg_wo_per_wk.reset_index()

# calculate avg hrs per WO by adding column that way you don't have to perform
# that calculation over and over again when you write the data
avg_wo_per_wk['HrsPerWO'] = avg_wo_per_wk['Est Hrs WO Calculated_v'] / \
                           avg_wo_per_wk['WOcount']

# drop 'Est Hrs WO Calculated_v' now that we're done with it
del avg_wo_per_wk['Est Hrs WO Calculated_v']

# round the count of WO's up to nearest whole number using "ceil"
avg_wo_per_wk['WOcount'] = avg_wo_per_wk['WOcount'].apply(math.ceil)
                                                      
# create a table of time deltas for the Priority Days.
wo_pri_days_raw = [(1, timedelta(4)), (2, timedelta(8)), (3, timedelta(14)), 
              (4, timedelta(30))]
wo_pri_days = pd.DataFrame(data = wo_pri_days_raw, columns=['WO Priority', 
                                                   'PriorityDays'])

# add wo_pri_days timedelta data to end of avg_wo_per_wk data.
# If WO priority = priority number then column timedelta equals said thing
avg_wo_per_wk = avg_wo_per_wk.merge(wo_pri_days, on='WO Priority', how='left')



'''
Create a dataframe for individual projected work orders
'''
# create the dataframe to write data to
x_wo_num = []
x_enter_date = []
x_due_date = []
x_crew = []
x_craft = []
x_priority = []

x_data_set = list(zip(x_wo_num,x_enter_date,x_due_date,x_crew,x_craft,
                     x_priority))
x_df = pd.DataFrame(data = x_data_set, columns=['WO Num', 'Enter Date', 
                                               'Due Date', 'Crew', 'Craft_v', 
                                               'WO Priority'])



'''
Create loop to write data
This loop goes through the avg_wo_per_wk data and for each row creates an
individual "dummy" or projected WO and creates data from that so there is
and dataframe of individual projected WO's as if they've already been created.
'''
# begin building data for projected WO's
today = pd.to_datetime('today')
# project two weeks worth of data from todays date
Week1 = today + timedelta(7)
Week2 = today + timedelta(14)

# assign all work orders to have an enter date of these mondays
FirstMonday = (Week1 - timedelta(days=Week1.weekday()))
SecondMonday = (Week2 - timedelta(days=Week2.weekday()))

i = 0 # row of data we're pulling
while i < len(avg_wo_per_wk):
    
    # select start and due dates
    Crew = avg_wo_per_wk.ix[i,'WO Crew']
    WO_Zone = avg_wo_per_wk.ix[i,'Zone']
    Craft_v = avg_wo_per_wk.ix[i,'Craft_v']
    WO_Priority = avg_wo_per_wk.ix[i,'WO Priority']
    WO_Count = avg_wo_per_wk.ix[i,'WOcount']
    HrsPerWO = avg_wo_per_wk.ix[i,'HrsPerWO']
    PriorityDays = avg_wo_per_wk.ix[i,'PriorityDays']
    WO_Num = 'Projected'    

    ii = 0 # iterate through number of estimated WO's
    while ii < WO_Count:
        aa = pd.DataFrame([[Crew, WO_Zone, Craft_v, WO_Priority, WO_Num, 
                            HrsPerWO, FirstMonday, FirstMonday+PriorityDays]], 
                            columns=['Crew', 'WO Zone', 'Craft_v', 
                                     'WO Priority', 'WO Num', 'HrsPerWO', 
                                     'Enter Date', 'Due Date'])
        x_df = x_df.append(aa, ignore_index=True)
        ii = ii + 1
        
    # stop counter
    i = i + 1

# copy the same average data except for the next week
week2_df = x_df.copy()

# same avg data, just need to create it for the next week, so should have an 
# enter date of the second Monday, and due dates will all be adjust one week
week2_df['Enter Date'] = week2_df['Enter Date'] + timedelta(7)
week2_df['Due Date'] = week2_df['Due Date'] + timedelta(7)
x_df = x_df.append(week2_df, ignore_index=True)

x_df.to_excel('Projected SR.xlsx', index=False)
print('Done')
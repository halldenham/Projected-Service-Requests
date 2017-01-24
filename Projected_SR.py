# -*- coding: utf-8 -*-
"""
Created on Mon Nov 28 15:57:18 2016

@author: dh1023
"""

# Import libraries
import pandas as pd
#import matplotlib
import numpy as np
#import sys
#import matplotlib
from datetime import timedelta
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
Merge/Join the Zone and Building Information to SR data and updated the Crews
to all be OPS x instead of ZONE x
'''
# change the name of the bldg number to make merge/join easier
bldg_df.columns = ['WO Building', 'Zone']

# make sure the building numbers are numbers, not text from WebI
sr_df['WO Building'] = pd.to_numeric(sr_df['WO Building'], errors='coerce')
bldg_df['WO Building'] = pd.to_numeric(bldg_df['WO Building'], errors='coerce')

# Merge/join data from Buldings so that the work orders have the correct 
# current zone/ops associated with them
sr_df = sr_df.merge(bldg_df, on='WO Building', how='left')


# list of the zones to change in to "OPS x" for following loop
zone_list = ['ZONE 1', 'ZONE 2', 'ZONE 4', 'ZONE 5']

# this loop takes WO's that currently have "ZONE x" and looks at what Ops area
# the building is currently in (which is the "Zone" column) and gives the WO
# that OPS as the crew instead of Zone.
i = 0
while i < len(sr_df):    
    # change any 'Zone 3' crews to 'Ops C'
    if sr_df.ix[i,'WO Crew'] == 'ZONE 3':
        sr_df.ix[i,'WO Crew'] = 'OPS C'
    # change any other "Zone x" to "Ops x"
    elif sr_df.ix[i,'WO Crew'] in zone_list:
        ops_letter = sr_df.ix[i,'Zone']
        # check to make sure ops_letter is actually a letter: if there's no
        # letter then there is no bldg associated with the WO
        if pd.isnull(ops_letter):
            new_ops = 'no_bldg'
        else:
            new_ops = 'OPS ' + ops_letter
        sr_df.ix[i,'WO Crew'] = new_ops

    i = i + 1


'''
Find the average per week of work orders and use as our projection
'''
# how many weeks worth of data do we have?
weeks_of_data = (sr_df['Enter Date'].max() - sr_df['Enter Date'].min()).days/7

# How many WO's were entered by priority, craft, & zone over the given period?
# Use ".size()" instead of .count() because size counts null values.
# Note: our historical average becomes our projected work.
previous_wo = sr_df.groupby(['WO Crew', 'Craft_v', 'WO Priority']).agg({
                            'WO Num': np.size, 'Est Hrs WO Calculated_v': 
                                np.sum})

# rename WO Num column to "WOcount" since it no longer is individual WO's
# but a summation of WO's
previous_wo = previous_wo.rename(columns = {'WO Num':'WOcount'})

# what does that average out to be per week?
avg_wo_per_wk = previous_wo / weeks_of_data

# reset index to actually create a dataframe
avg_wo_per_wk = avg_wo_per_wk.reset_index()

# now remove any WO's that don't have a building...no way to account for those
avg_wo_per_wk = avg_wo_per_wk[avg_wo_per_wk['WO Crew'] != 'no_bldg']

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
# create a dataframe for Excel data
x_df = pd.DataFrame(columns=['WO Num', 'Enter Date', 'Due Date', 'Crew', 
                             'Craft_v', 'WO Priority'])


'''
Create loop to write data
This loop goes through the avg_wo_per_wk data and for each row creates an
individual "dummy" or projected WO and creates data from that so there is
and dataframe of individual projected WO's as if they've already been created.
'''

today = pd.to_datetime('today')
# project two weeks worth of data from todays date
week_1 = today + timedelta(7)
week_2 = today + timedelta(14)

# assign all work orders to have an enter date of these mondays
first_monday = (week_1 - timedelta(days=week_1.weekday()))
second_monday = (week_2 - timedelta(days=week_2.weekday()))

i = 0 # row of data we're pulling
while i < len(avg_wo_per_wk):
    
    # select relevant data for when the days get split up
    crew = avg_wo_per_wk.ix[i,'WO Crew']
    craft_v = avg_wo_per_wk.ix[i,'Craft_v']
    wo_priority = avg_wo_per_wk.ix[i,'WO Priority']
    wo_count = avg_wo_per_wk.ix[i,'WOcount']
    hrs_per_wo = avg_wo_per_wk.ix[i,'HrsPerWO']
    pri_days = avg_wo_per_wk.ix[i,'PriorityDays']
    wo_num = 'Projected'    

    ii = 0 # iterate through number of estimated WO's
    while ii < wo_count:
        aa = pd.DataFrame([[crew, craft_v, wo_priority, wo_num, 
                            hrs_per_wo, first_monday, first_monday+pri_days]], 
                            columns=['Crew', 'Craft_v', 
                                     'WO Priority', 'WO Num', 'HrsPerWO', 
                                     'Enter Date', 'Due Date'])
        x_df = x_df.append(aa, ignore_index=True)
        ii = ii + 1
        
    # stop counter
    i = i + 1

# copy the same average data except for the next week
week_2_df = x_df.copy()

# same avg data, just need to create it for the next week, so should have an 
# enter date of the second Monday, and due dates will all be adjust one week
week_2_df['Enter Date'] = week_2_df['Enter Date'] + timedelta(7)
week_2_df['Due Date'] = week_2_df['Due Date'] + timedelta(7)
x_df = x_df.append(week_2_df, ignore_index=True)

x_df.to_excel('Projected SR.xlsx', index=False)
print('Done')
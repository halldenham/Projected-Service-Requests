# -*- coding: utf-8 -*-
"""
Created on Mon Nov 28 15:57:18 2016

@author: dh1023
"""

# Import libraries
import pandas as pd
import numpy as np
from datetime import timedelta
import math
import tkinter as tk


'''
function to convert 'zone x' to 'ops x'
'''
def zone_ops(wo_dataframe):
    # this function converts a list of crews from the old 'Zone X' to 'Ops X'    
    
    zone_list = ['ZONE 1', 'ZONE 2', 'ZONE 4', 'ZONE 5']
    # list of the zones to change in to 'OPS x'
    
    if wo_dataframe['WO Crew'] == 'ZONE 3':
        # change any 'Zone 3' crews to 'Ops C'
        new_crew = 'OPS C' 
    elif wo_dataframe['WO Crew'] in zone_list:
        # change the work order crew from 'ZONE X' to 'OPS X' where 'X' is the
        # current zone for the building with which the old work order is 
        # associated
        
        if pd.isnull(wo_dataframe['Zone']):
            # if it is a Zone crew, but the work order does not have a 
            # building associated with it, then error
            new_crew = 'error_no_bldg'
        else:
            new_crew = 'OPS ' + wo_dataframe['Zone']
    else:
        # if the crew is not 'zone x' then leave crew as-is
        # this takes care of 'MCORD' and any crews that are already 'ops x'
        new_crew = wo_dataframe['WO Crew']
    return new_crew;
        


'''
Import the data

At some point in future, create a dialog box to prompt user to select file
root = tk.Tk()
root.withdraw()
file_path = tk.filedialog.askopenfilename()
'''
# Location of file with previous SR data
# This data can have any length of date range and should include all WO's
# This range of data usually mirrors the time period for which it is seeking
# to predict: i.e. if the next two weeks are Dec 1-14, 2016, then the data
# should cover Dec 1-14, 2015 and potentially some buffer on either side.
# This is the raw data used to predict future SR.
sr_excel = r'C:\Users\dh1023\Desktop\Python\Zone_Manager_Report_SR_Project' \
         'ed_-_DGH_Master.xlsx'

# Create dataframe from previous SR data
sr_df = pd.read_excel(sr_excel)

# Location of file with data of what buildings belong to what zones
# Note that first column should be "Bldg/ Land Entity" which is the building 
# number and second column should be "Zone" 
bldg_excel = r'C:\Users\dh1023\Desktop\Python\Reference Files or ' \
              'Standards\FMS61100_-_Active_Building_or_Land_Entity' \
              '_Listing.xlsx'

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

# convert the "Zone" crew to "Ops" crew using function
sr_df['WO Crew'] = sr_df.apply(zone_ops, axis=1)



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



'''
Cleaning up the data
'''
# remove any WO's that don't have a building...no way to account for those
avg_wo_per_wk = avg_wo_per_wk[avg_wo_per_wk['WO Crew'] != 'no_bldg']

# calculate avg hrs per WO by adding column that way you don't have to perform
# that calculation over and over again when you write the data
avg_wo_per_wk['HrsPerWO'] = avg_wo_per_wk['Est Hrs WO Calculated_v'] / \
                           avg_wo_per_wk['WOcount']

# drop 'Est Hrs WO Calculated_v' now that we're done with it
del avg_wo_per_wk['Est Hrs WO Calculated_v']

# round the count of WO's up to nearest whole number using "ceil"
avg_wo_per_wk['WOcount'] = avg_wo_per_wk['WOcount'].apply(math.ceil)
                                                      

# add wo_pri_days timedelta data to end of avg_wo_per_wk data.
# create a table of time deltas for the Priority Days.
wo_pri_days_raw = [(1, timedelta(4)), (2, timedelta(8)), (3, timedelta(14)), 
              (4, timedelta(30))]
wo_pri_days = pd.DataFrame(data = wo_pri_days_raw, columns=['WO Priority', 
                                                   'PriorityDays'])

# If WO priority = priority number then column timedelta equals said thing
avg_wo_per_wk = avg_wo_per_wk.merge(wo_pri_days, on='WO Priority', how='left')

# change priorities from 1 to 101, 2 to 102, 3 to 103, 4 to 104 for use in 
# Tableau.
avg_wo_per_wk['WO Priority'] = avg_wo_per_wk['WO Priority']+100


# give work orders a start and due date of the upcoming week and attached to 
# avg_wo_per_wk dataframe
today = pd.to_datetime('today')
# project two weeks worth of data from todays date
week_1 = today + timedelta(7)
week_2 = today + timedelta(14)

# assign all work orders to have an enter date of this monday
first_monday = (week_1 - timedelta(days=week_1.weekday()))

avg_wo_per_wk['Enter Date'] = first_monday
avg_wo_per_wk['Due Date'] = first_monday + avg_wo_per_wk['PriorityDays']



'''
Create loop to write data
This loop goes through the avg_wo_per_wk data and for each row creates an
individual "dummy" or projected WO and creates data from that so there is
and dataframe of individual projected WO's as if they've already been created.
'''
x_df = pd.DataFrame(columns=[])
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

second_monday = (week_2 - timedelta(days=week_2.weekday()))

# same avg data, just need to create it for the next week, so should have an 
# enter date of the second Monday, and due dates will all be adjust one week
week_2_df['Enter Date'] = week_2_df['Enter Date'] + timedelta(7)
week_2_df['Due Date'] = week_2_df['Due Date'] + timedelta(7)
x_df = x_df.append(week_2_df, ignore_index=True)

x_df.to_excel('Projected SR.xlsx', index=False)
print('Done')



'''
this is an alternate function way of the code, but I haven't
quite figured it out...this works but it's slower than the existing code as
of 3/24/17 so I'm not going to use it...but I'll let it hang out here
'''
#def proj_wo(old_sr):
#    # something here
#    # for each row of data, copy it n times into a dataframe where n = WOcount
#  
#    df = pd.DataFrame(columns=[])
#    # create empty dataframe to append calculated data to
#    
#    for i in range(old_sr['WOcount']):
#        df = df.append(old_sr, ignore_index=True)
#    
#    df['WO Num'] = 'Projected'
#    return df;
#
#x_df = pd.DataFrame(columns=[])
#i = 0 # row of data we're pulling
#
#
#while i < len(avg_wo_per_wk):
#    aa = proj_wo(avg_wo_per_wk.ix[i,:])
#    x_df = x_df.append(aa)
#    i = i + 1

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
# Import all open WO data
# Criteria for WebI Report:
# Req Type = "PM; SR"
# WO Status = "Assigned; Open"
# Crews = "All"
KPIdata = r'C:\Users\dh1023.AD\Desktop\Python\Zone_Manager_Report_Open_and_' \
          'Closed_WO_-_DGH_Master.xlsx'
KPI_df = pd.read_excel(KPIdata)

# Location of file with previous SR data
# This data can have any length of date range and should include all WO's
# This range of data usually mirrors the time period for which it is seeking
# to predict: i.e. if the next two weeks are Dec 1-14, 2016, then the data
# should cover Dec 1-14, 2015 and potentially some buffer on either side
SRdata = r'C:\Users\dh1023.AD\Desktop\Python\Zone_Manager_Report_SR_Project' \
         'ed_-_DGH_Master.xlsx'

# Create dataframe from previous SR data
SR_df = pd.read_excel(SRdata)

# Location of file with data of what buildings belong to what zones
# Note that first column should be "Bldg/ Land Entity" which is the building 
# number and second column should be "Zone" 
BuildingData = r'C:\Users\dh1023.AD\Desktop\Python\FMS61100_-_Active_Build' \
               'ing_or_Land_Entity_Listing.xlsx'

# Create dataframe from building data
Bldg_df = pd.read_excel(BuildingData)





'''
Merge/Join the Zone and Building Information to SR data
'''
# change the name of the bldg number to make merge/join easier
Bldg_df.columns = ['WO Building', 'Zone']

# make sure the building numbers are numbers, not text from WebI
SR_df['WO Building'] = pd.to_numeric(SR_df['WO Building'], errors='coerce')
Bldg_df['WO Building'] = pd.to_numeric(Bldg_df['WO Building'], errors='coerce')

# Merge/join data from Buldings so that the work orders have the correct 
# current zone/ops assigned to them
SR_df = SR_df.merge(Bldg_df, on='WO Building', how='left')






'''
Find the average per week of work orders and use as our projection
'''
# how many weeks worth of data do we have?
weeksOfData = (SR_df['Enter Date'].max() - SR_df['Enter Date'].min()).days/7

# how many WO's were entered by priority and craft over the given period?
# what is the sum of esimated hours of the given period?
# use ".size()" instead of .count() because size counts null values
# our historical average becomes our projected work

previousWO = SR_df.groupby(['WO Crew','Zone','Craft_v','WO Priority']).agg({
                            'WO Num': np.size, 'Est Hrs WO Calculated_v': 
                                np.sum})

# rename WO Num column to "WOcount" since it no longer is individual WO's
# but a summation of WO's
previousWO = previousWO.rename(columns = {'WO Num':'WOcount'})

# what does that average out to be per week?
avgWOperWeek = previousWO / weeksOfData

# reset index to actually create a dataframe
avgWOperWeek = avgWOperWeek.reset_index()

# calculate avg hrs per WO by adding column that way you don't have to perform
# that calculation over and over again when you write the data
avgWOperWeek['HrsPerWO'] = avgWOperWeek['Est Hrs WO Calculated_v'] / \
                           avgWOperWeek['WOcount']

# drop 'Est Hrs WO Calculated_v' now that we're done with it
del avgWOperWeek['Est Hrs WO Calculated_v']

# round count of WO's up to nearest whole number using "ceil"
avgWOperWeek['WOcount'] = avgWOperWeek['WOcount'].apply(math.ceil)
                                                      
# add column for days to add to enter date by creating a table for priority 
# day time deltas for reference
WOPDaysRaw = [(1, timedelta(4)), (2, timedelta(8)), (3, timedelta(14)), 
              (4, timedelta(30))]
WOPDays = pd.DataFrame(data = WOPDaysRaw, columns=['WO Priority', 
                                                   'PriorityDays'])
# add WOPDays timedelta data to end of avgWOperWeek data
# if WO priority = number then column timedelta equals said thing
avgWOperWeek = avgWOperWeek.merge(WOPDays, on='WO Priority', how='left')



'''
Create a dataframe for individual projected work orders
'''
# create the dataframe to write data to
x_WOnum = []
x_EnterDate = []
x_DueDate = []
x_Crew = []
x_Craft = []
x_Priority = []

x_dataSet = list(zip(x_WOnum,x_EnterDate,x_DueDate,x_Crew,x_Craft,
                     x_Priority))
x_df = pd.DataFrame(data = x_dataSet, columns=['WO Num', 'Enter Date', 
                                               'Due Date', 'Crew', 'Craft_v', 
                                               'WO Priority'])



'''
Create loop to write data
This loop goes through the avgWOperWeek data and for each row creates an
individual "dummy" or projected WO and creates a data from so that there is
and dataframe of individual projected WO's as if they've already been created.
'''
# begin building data for projected WO's
today = pd.to_datetime('today')
Week1 = today + timedelta(7)
Week2 = today + timedelta(14)

# assign all work orders to have an enter date of these mondays
FirstMonday = (Week1 - timedelta(days=Week1.weekday()))
SecondMonday = (Week2 - timedelta(days=Week2.weekday()))

i = 0 # row of data we're pulling
while i < len(avgWOperWeek):
    
    # select start and due dates
    Crew = avgWOperWeek.ix[i,'WO Crew']
    WO_Zone = avgWOperWeek.ix[i,'Zone']
    Craft_v = avgWOperWeek.ix[i,'Craft_v']
    WO_Priority = avgWOperWeek.ix[i,'WO Priority']
    WO_Count = avgWOperWeek.ix[i,'WOcount']
    HrsPerWO = avgWOperWeek.ix[i,'HrsPerWO']
    PriorityDays = avgWOperWeek.ix[i,'PriorityDays']
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
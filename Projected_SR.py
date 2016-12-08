# -*- coding: utf-8 -*-
"""
Created on Mon Nov 28 15:57:18 2016

@author: dh1023
"""

# For a given data set, find how many weeks there are then find avg by 
# priority and craft and add est hours to a dataframe

# Import libraries
import pandas as pd
import matplotlib
import numpy as np
import sys
import matplotlib
from datetime import datetime, timedelta, date
import datetime as dt

'''
Import the data
'''
# Import actual open WO data
KPIdata = r'C:\Users\dh1023.AD\Desktop\Tableau files\Zone_Manager_Report_Open_and_Closed_WO_-_DGH_Master.xlsx'
KPI_df = pd.read_excel(KPIdata)

# Location of file with previous SR data
SRdata = r'C:\Users\dh1023.AD\Desktop\Python\Zone_Manager_Report_SR_Projected_-_DGH_Master.xlsx'
# Create dataframe from previous SR data
SR_df = pd.read_excel(SRdata)

# Location of file with data of what buildings belong to what zones
# Note that first column should be "Bldg/ Land Entity" which is the building number
# and second column should be "Zone" 
BuildingData = r'C:\Users\dh1023.AD\Desktop\Python\FMS61100_-_Active_Building_or_Land_Entity_Listing.xlsx'
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

# Merge/join data from Buldings so that the work orders have the correct current
# zone/ops assigned to them
SR_df = SR_df.merge(Bldg_df, on='WO Building', how='left')



'''
Find the average per week of work orders
'''

# how many weeks worth of data do we have?
weeksOfData = (SR_df['Enter_Date'].max() - SR_df['Enter_Date'].min()).days/7

# how many WO's were entered by priority and craft over the given period?
# what is the sum of esimated hours of the given period?
# use ".size()" instead of .count() because size counts null values
projectedWO = SR_df.groupby(['Craft_v','WO Priority']).agg({'WO_Num': np.size, 'Est Hrs WO Calculated_v': np.sum})

# what does that average out to be per week?
avgWOperWeek = projectedWO / weeksOfData


#!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
# when we exceed the previous years number of work orders, we get negative numbers
# how should we adjust our projections when we're exceeding previous avgs?
# solve later maybe?
'''
Look at existing work orders over the next two weeks and subtract those from projections
'''
today = pd.to_datetime('today')
twoWeeks = today + timedelta(14)

# find all WO's between today and 2 weeks to remove from projected 2 weeks
ActualWOtwoWeeks = KPI_df[(KPI_df['Comp Due Date_v'] > today) & (KPI_df['Comp Due Date_v'] < twoWeeks)].groupby(['Craft_v','WO Priority']).agg({'WO_Num': np.size, 'Est Hrs WO Calculated_v': np.sum})

# we're looking at the actual work orders over the next two weeks, so we need
# to multiple the avg per week by two to get projected work orders for the 
# next two weeks
ProjectedWork = (avgWOperWeek*2) - ActualWOtwoWeeks
ProjectedWork = ProjectedWork.reset_index()

# create data frame to perform calculations


'''
Create a data set with individual projected work orders
'''
# create the dataframe to write data to
x_Date = []
x_Hrs = []
x_WOnum = []
x_SRnum = []
x_Crew = []
x_Craft = []
x_Priority = []
x_Bldg = []
x_BldgZone = []
x_EnterDate = []
x_DueDate = []
x_dataSet = list(zip(x_Date,x_Hrs,x_WOnum,x_SRnum,x_Crew,x_Craft,x_Priority,x_Bldg,x_BldgZone,x_EnterDate,x_DueDate))
x_df = pd.DataFrame(data = x_dataSet, columns=['Date', 'Est Hrs', 'WO Num', 'SR Num', 'Crew', 'Craft', 'Prriority', 'Building', 'Building Zone', 'Enter Date', 'Due Date'])


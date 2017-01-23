# -*- coding: utf-8 -*-
"""
Created on Fri Nov  4 10:05:13 2016

@author: dh1023
"""

# Import libraries
import pandas as pd
import matplotlib
import numpy as np
import sys
import matplotlib
from datetime import datetime, timedelta, date

# Location of file with data
# Criteria for Report:
# Req Type = "PM; SR"
# WO Status = "Assigned; Open"
# Crews = "All"

location = r'C:\Users\dh1023.AD\Desktop\Python\Zone_Manager_Report_Open_and' \
            '_Closed_WO_-_DGH_Master.xlsx'


# alternate location for Projected SR
#location = r'C:\Users\dh1023.AD\Desktop\Python\Projected SR.xlsx'            

            
# Parse a specific sheet
df = pd.read_excel(location)

# create the dataframe to write data to
x_Date = []
x_Hrs = []
x_wo_num = []
x_SRnum = []
x_Crew = []
x_Craft = []
x_Priority = []
x_dataSet = list(zip(x_Date,x_Hrs,x_wo_num,x_SRnum,x_Crew,x_Craft,x_Priority))
x_df = pd.DataFrame(data = x_dataSet, columns=['Date', 'Est Hrs', 'WO Num', 
                                               'SR Num'])


today = pd.to_datetime(date.today())
# begin loop!!!!!!!!!!!!!
i = 0 # row of data we're pulling

while i < len(df):
    
    # select start and due dates
    wo_num = df.iloc[i,0]
    avg_hrs = df.iloc[i,1]
    enter_date = df.iloc[i,2]
    due_date = df.iloc[i,3]    
    sr_num = df.iloc[i,4]
    crew = df.iloc[i,5]
    craft = df.iloc[i,6]
    priority = df.iloc[i,7]

    # select the earliest date available to work
    if due_date <= today:
        # append only 1 row of data if it's already past due
        due_date_mon = (due_date - timedelta(days=due_date.weekday()))
        aa = pd.DataFrame([[due_date_mon, avg_hrs, wo_num, sr_num, crew, craft, 
                            priority]], columns=['Date', 'Est Hrs', 'WO Num', 
                                                 'SR Num', 'crew', 'craft', 
                                                 'priority'])
        x_df = x_df.append(aa, ignore_index=True)
        
    else:
        
        # convert start and due dates to the Monday of the week they fall in
        
        # if the WO "Start Date" is prior to today, and is due after today, 
        # then we need to select the available hours for only the weeks we 
        # still have available to work. In other words, don't put estimated 
        # time to prior weeks if the WO is still open and not overdue
        today_mon = (today - timedelta(days=today.weekday()))
        original_mon = (enter_date - timedelta(days=enter_date.weekday()))
        monday1 = max(today_mon,original_mon)
        monday2 = (due_date - timedelta(days=due_date.weekday()))
        
        # list of the mondays available for the WO to be completed
        df_weeks_between = pd.date_range(monday1, monday2, freq='W-Mon')

        # different information to write back to our data
        weeks_between = 1 + (monday2 - monday1).days / 7
        
        # find average hours per week
        avg_hrs_per_week = avg_hrs / weeks_between        
    
        # now write Avg Hrs per week to each of the mondays between monday1 
        # and monday 2
        # write those into a dataframe and then append it to the original data
        for x in df_weeks_between: 
            aa = pd.DataFrame([[x, avg_hrs_per_week, wo_num, sr_num, crew, craft, 
                                priority]], columns=['Date', 'Est Hrs', 
                                                     'WO Num', 'SR Num', 
                                                     'crew', 'craft', 
                                                     'priority'])
            x_df = x_df.append(aa, ignore_index=True)

    # loop through all the rows    
    i = i + 1

# export results to Excel   
x_df.to_excel('Estimated Hrs per week.xlsx', index=False)


# alternate location for taking the SR projected
#x_df.to_excel('Estimated Hrs per week SR Proj.xlsx', index=False)


print('Done')

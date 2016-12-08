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
Location = r'C:\Users\dh1023.AD\Desktop\Python\Zone_Manager_Report_Open_and_Closed_WO_-_DGH_Master.xlsx'

# Parse a specific sheet
df = pd.read_excel(Location)

# create the dataframe to write data to
x_Date = []
x_Hrs = []
x_WOnum = []
x_SRnum = []
x_Crew = []
x_Craft = []
x_Priority = []
x_dataSet = list(zip(x_Date,x_Hrs,x_WOnum,x_SRnum,x_Crew,x_Craft,x_Priority))
x_df = pd.DataFrame(data = x_dataSet, columns=['Date', 'Est Hrs', 'WO Num', 'SR Num'])


today = pd.to_datetime(date.today())
# begin loop!!!!!!!!!!!!!
i = 0 # row of data we're pulling

while i < len(df):
    
    # select start and due dates
    dueDate = df.iloc[i,3]    
    enterDate = df.iloc[i,2]
    avgHrs = df.iloc[i,1]
    WOnum = df.iloc[i,0]
    SRnum = df.iloc[i,4]
    Crew = df.iloc[i,5]
    Craft = df.iloc[i,6]
    Priority = df.iloc[i,7]

    # select the earliest date available to work
    if dueDate <= today:
        # append only 1 row of data if it's already past due
        dueDateMonday = (dueDate - timedelta(days=dueDate.weekday()))
        aa = pd.DataFrame([[dueDateMonday, avgHrs, WOnum, SRnum, Crew, Craft, Priority]], columns=['Date', 'Est Hrs', 'WO Num', 'SR Num', 'Crew', 'Craft', 'Priority'])
        x_df = x_df.append(aa, ignore_index=True)
        
    else:
        
        # convert start and due dates to the Monday of the week they fall in
        
        # if the WO "Start Date" is prior to today, and is due after today, then 
        # we need to select the available hours for only the weeks we still have
        # available to work. In other words, don't put estimated time to prior
        # weeks if the WO is still open and not overdue
        todayMonday = (today - timedelta(days=today.weekday()))
        originalMonday = (enterDate - timedelta(days=enterDate.weekday()))
        monday1 = max(todayMonday,originalMonday)
        monday2 = (dueDate - timedelta(days=dueDate.weekday()))
        
        # list of the mondays available for the WO to be completed
        dfWeeksBetween = pd.date_range(monday1, monday2, freq='W-Mon')

        # different information to write back to our data
        weeksBetween = 1 + (monday2 - monday1).days / 7
        
        # find average hours per week
        avgHrsPerWeek = avgHrs / weeksBetween        
    
        # now write Avg Hrs per week to each of the mondays between monday1 and monday 2
        # write those into a dataframe and then append it to the original data
        for x in dfWeeksBetween: 
            aa = pd.DataFrame([[x, avgHrsPerWeek, WOnum, SRnum, Crew, Craft, Priority]], columns=['Date', 'Est Hrs', 'WO Num', 'SR Num', 'Crew', 'Craft', 'Priority'])
            x_df = x_df.append(aa, ignore_index=True)

    # loop through all the rows    
    i = i + 1

# export results to Excel   
x_df.to_excel('Estimated Hrs per week.xlsx', index=False)
print('Done')

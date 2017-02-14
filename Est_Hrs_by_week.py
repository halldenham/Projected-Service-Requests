# -*- coding: utf-8 -*-
"""
Created on Fri Nov  4 10:05:13 2016

@author: dh1023
"""

# Import libraries
import pandas as pd
from datetime import timedelta, date



"""
import excel worksheets
"""

# Location of file with data
# Criteria for Report:
# Req Type = "PM; SR"
# WO Status = "Assigned; Open"
# Crews = "All"
wo_excel = r'C:\Users\dh1023.AD\Desktop\Python\Zone_Manager_Report_Open_and' \
            '_Closed_WO_-_DGH_Master.xlsx'
# create a dataframe
open_assigned_df = pd.read_excel(wo_excel)

# location for Projected SR for other python code
proj_sr_excel = r'C:\Users\dh1023.AD\Desktop\Python\Projected SR.xlsx'
# create a dataframe
proj_sr_df = pd.read_excel(proj_sr_excel)



"""
Join the projected SR with the actual open WO's
"""
# ensure the two dataframes have the same column names
proj_sr_df = proj_sr_df.rename(columns={'Due Date': 'Comp Due Date_v', \
                                        'HrsPerWO': 'Est Hrs WO Calculated_v'\
                                        , 'Crew': 'WO Crew'})

# add the projected WO's to the actual open and assigned WO's
wo_df = open_assigned_df.append(proj_sr_df, ignore_index=True)





"""
Find the estimated hours
"""
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

while i < len(wo_df):
    
    # select start and due dates
    wo_num = wo_df.ix[i,'WO Num']
    avg_hrs = wo_df.ix[i,'Est Hrs WO Calculated_v']
    enter_date = wo_df.ix[i,'Enter Date']
    due_date = wo_df.ix[i,'Comp Due Date_v']    
    req_num = wo_df.ix[i,'Req Number']
    crew = wo_df.ix[i,'WO Crew']
    craft = wo_df.ix[i,'Craft_v']
    priority = wo_df.ix[i,'WO Priority']

    # select the earliest date available to work
    if due_date <= today:
        # append only 1 row of data if it's already past due
        due_date_mon = (due_date - timedelta(days=due_date.weekday()))
        aa = pd.DataFrame([[due_date_mon, avg_hrs, wo_num, req_num, crew, craft, 
                            priority]], columns=['Date', 'Est Hrs', 'WO Num', 
                                                 'SR Num', 'Crew', 'Craft', 
                                                 'Priority'])
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
            aa = pd.DataFrame([[x, avg_hrs_per_week, wo_num, req_num, crew, 
                                craft, priority]], columns=['Date', 'Est Hrs', 
                                                     'WO Num', 'SR Num', 
                                                     'Crew', 'Craft', 
                                                     'Priority'])
            x_df = x_df.append(aa, ignore_index=True)

    # loop through all the rows    
    i = i + 1

"""
add on the available hours to data from reference files
"""
# import data from reference files
avail_hrs = r'C:\Users\dh1023.AD\Desktop\Python\Reference Files or ' \
            'Standards\Ops Available Hrs.xlsx'
# create a dataframe
avail_hrs_df = pd.read_excel(avail_hrs)

# update the empty "date" field to a couple days in future so it shows up in
# Tableau report fields when filtering for data in a two week range
avail_hrs_df['Date'] = today + timedelta(2)

# add that data to the end of the estimated hours
x_df = x_df.append(avail_hrs_df, ignore_index=True)
    
# export results to Excel   
x_df.to_excel('Estimated Hrs per week.xlsx', index=False)


print('Done')

# -*- coding: utf-8 -*-
"""
Created on Wed Jun  7 12:51:11 2017

@author: dh1023
"""

import pandas as pd

# create a loop that selects all relevant wo's based on crew, then randomizes
# the order, then select the the first x random values needed
# print that list to an excel file with relevant data


# Location of file with WO data. Include PM and SR that were completed within
# the previous week.
wo_excel = r'C:\Users\dh1023\Desktop\Python\Generic_WO_Report_-_QC_and_QA_' \
            'Report.xlsx'
# Create dataframe from previous SR data
wo_df = pd.read_excel(wo_excel)


# create a df of crews and how many random QC/QA work orders are needed
ops_list = ['OPS A', 'OPS B', 'OPS C', 'OPS D']
qc_list = [11, 20, 2, 24]
qa_list = [8, 16, 2, 19]
ops_qc_qa_zip = list(zip(ops_list, qc_list, qa_list))
ops_qc_qa = pd.DataFrame(data = ops_qc_qa_zip, columns=['Crew', 'QC', 'QA'])


i = 0 # this is used to index crews and qa/qc values or data
while i<len(ops_qc_qa):
    # select random data with current crew (i value)
    qc_wo = wo_df.loc[wo_df['Crew']==ops_list[i]].sample(qc_list[i])
    qa_wo = wo_df.loc[wo_df['Crew']==ops_list[i]].sample(qa_list[i])
    
    # add columns to specify if QC or QA
    qc_wo['QC or QA'] = 'QC'
    qa_wo['QC or QA'] = 'QA'
    
    # append dataframes together
    qc_wo = qc_wo.append(qa_wo, ignore_index=True)
    
    # export to an excel file
    qc_wo.to_excel(r'QC and QA work orders for '+ops_list[i]+'.xlsx', 
                   index=False)
    
    print('Completed', ops_list[i])
    i +=1
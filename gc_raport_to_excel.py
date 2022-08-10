# import libraries
from pdfminer.high_level import extract_text
import pandas as pd
import numpy as np
import re
import os

######################################### FUNCTIONS ###########################################
def make_table(text_data):
    '''
    Takes standart GC-MS report and extracts the valuable columns, namely:
    1. Retention time  as  RetTime [min]
    2. w/w amount in % as  Amount_%
    3. Substance name  as  Name
    ''' 
    
    # extract the columns from the text file
    data = re.findall(r'(\d+\.\d{3})\s+(\w{2}\s+[A-Z]?\s+\d+)\s+(\d+\.\d+[e\-1]*)\s+(\d+\.\d+[e\-1]*)\s+(\d+\.\d+)\s+(\w+\s\w+\s?\w*|\?)',
                      text_data)
    
    # define column names
    col_names = 'RetTime [min],Type_ISTD used,Area [pA*s],Amt/Area_ratio,Amount_%,Name'
    col_names = col_names.split(',')
    
    # create temporary dataframe: tmp_df
    tmp_df = pd.DataFrame(np.array(data), columns=col_names)
    
    # define columns to be dropped and drop them
    cols_to_drop = tmp_df.columns[[1, 2, 3]]
    tmp_df.drop(cols_to_drop, axis=1, inplace=True)
    
    # change the datatype of numeric columns from object to float
    tmp_df.iloc[:, 0] = tmp_df.iloc[:, 0].astype(float)
    tmp_df.iloc[:, 1] = tmp_df.iloc[:, 1].astype(float)
    
    # return the dataframe
    return tmp_df

def file_number_to_loc(file_number):
    '''
    Takes the number|name of file in format xxx.xx, as this is the standard format 
    for MS we are using and returns full localization.
    '''
    _list_of_loc = [main_path + '\\' + name for name in dict_files[file_number]]
    return _list_of_loc

######################################### MAIN CODE ###########################################
# get groupped and name of final excel file
groupped = True
excel_name = 'GC-MS.xlsx'
excel_path = r'path_to_files'

# define path and create an array of file names
main_path = r'path_to_files'
files = os.listdir(main_path)

# flatten the array to 1d
files = np.array(files).flatten()

# drop the end strings leaving only numbers and get the unique values
number_files = np.unique([re.findall(r'\d+\.\d+', name)[0] for name in files])

# create a dictionary to store the files {unique_number:[file_name_1, ..., file_name_N]}
dict_files = {}
for num in number_files:
    
    # create a list of True/False as mask for subsetting 
    bool_list = [re.match(rf'{num}', f) != None for f in files]
    
    # subset above mask
    files_to_dict = files[bool_list]
    
    # if unique_number not in keys then add and map according list of names 
    if num not in dict_files.keys():
        dict_files[num] = files_to_dict

# save to excel file, every group of files to given sheet
with pd.ExcelWriter(excel_path + '\\' + excel_name, mode="w", engine="openpyxl", if_sheet_exists='new') as writer:
    
    for file_number in dict_files.keys():
        
        # create a temporary dataframe to store the results
        tmp_df = pd.DataFrame()
        
        for file_loc in file_number_to_loc(file_number):
            
            # append new entries to the df
            tmp_df = tmp_df.append(
                make_table(
                    extract_text(
                        file_loc
                    )
                )
            )
            
            # sort and reset the index
            tmp_df = tmp_df.sort_values(by=tmp_df.columns[0]).reset_index(drop=True)
        
        # if groupped is true group the values by the name of substance for mean values
        if groupped == True:
            tmp_df = tmp_df.groupby(tmp_df.columns[-1]).mean()
        
        # store in a excel file
        tmp_df.to_excel(writer, sheet_name=f'{file_number}', startrow=1, startcol=1)

print('Done!')
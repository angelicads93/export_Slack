import numpy as np
import pandas as pd

from inputs import missing_value

#class CleanDF:
#    def __init__(self):
#        pass
    
def replace_empty_space(df, column):
    """Function to replace empty spaces "" with the string missing_value for a given column"""
    for i in range(len(df)):
        if df.at[i,column] == "":
            df.at[i,column] = missing_value  
            
def replace_NaN(df, column):
    """Function to replace missing values with the string missing_value for a given column """
    df[column] = df[column].fillna(missing_value)    
    
def handle_missing_values(df):
    """Function that replaces missing values in all the columns of the df"""
    df = df.replace(pd.NaT, missing_value)
    df = df.replace(np.nan, missing_value) 
    df = df.fillna(missing_value)
    return df
    
def drop_rows(df, indices):
    try:
        df = df.drop(indices,axis=0)
    except:
        pass  
    return df

def reset_indices(df):
    df.index = np.arange(0,len(df),1)
    return df        

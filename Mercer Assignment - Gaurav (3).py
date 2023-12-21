#!/usr/bin/env python
# coding: utf-8

# Importing all libraries

# In[1]:


# Importing libraries

import numpy as np
import math as m
import pandas as pd
from numpy.linalg import multi_dot
pd.options.display.float_format = '{:.2%}'.format

# Ignore warnings
import warnings
warnings.filterwarnings('ignore')

# Input file path

file =(r'C:\Users\Gaurav\Desktop\Python_Assignment.xlsx')


# Input Dataframes for CMA sheet

# In[2]:


#Reading raw data from spreadsheet

def CMA_df1(file):
    
    CMA_df1 = pd.read_excel(file, sheet_name=2, usecols="B:E,G:M", header=1, index_col=0, nrows=115)
    
    return CMA_df1


# CMA sheet correlation dataframe

# In[3]:


def CMA_df2(file):

    CMA_df2 = pd.read_excel(file, sheet_name=2, usecols="N:DZ", header=1, index_col=0, nrows=116)
    CMA_df2.iloc[-1:] = CMA_df2.iloc[-1:].shift(periods=1, axis=1)
    CMA_df2.iloc[:,0] = CMA_df2.iloc[:,0].shift(1)
    CMA_df2.columns = CMA_df2.iloc[0]
    CMA_df2 = CMA_df2.iloc[1:]
    CMA_df2.set_index(CMA_df2.iloc[:,0].name, inplace= True)

    return CMA_df2


# Data manipulation in the intermediate Setup sheet (Risk/Return profile for various asset classes)

# In[4]:


def Setup_df1(file):
    
    X = CMA_df1(file)

    Setup_df1 = pd.read_excel(file, sheet_name=1, usecols="B:C", header=1, index_col=1, nrows=28)

    #Active Returns
    Setup_df1['Active Risk'] = Setup_df1.index.map(X["TE"])
    Setup_df1['Info Ratio'] = Setup_df1.index.map(X["IR (A Rated)"])
    Setup_df1['Active Return'] = Setup_df1.iloc[:,1:3].prod(axis=1)
    Setup_df1['Fee'] = pd.read_excel(file, sheet_name=1, usecols="C,G", header=1, index_col=0, nrows=28)
    Setup_df1['#Mgrs'] = pd.read_excel(file, sheet_name=1, usecols="C,H", header=1, index_col=0, nrows=28)

    #Passive Returns
    Setup_df1['Arithmetic Passive'] = Setup_df1.index.map(X["Expected  Annual Return"])
    Setup_df1['Volatility - SD Passive'] = Setup_df1.index.map(X["Annual Standard Deviation"])
    Setup_df1['Geometric Passive'] = np.exp(np.log(1+ Setup_df1.iloc[:,6])-(np.log(1+((Setup_df1.iloc[:,7]**2)/((1+Setup_df1.iloc[:,6])**2))))/2)-1

    #Total Returns
    Setup_df1['Arithmetic Total'] = Setup_df1['Arithmetic Passive'] + Setup_df1['Active Return']
    Setup_df1['Volatility - SD Total'] = np.sqrt(Setup_df1['Volatility - SD Passive']**2 + Setup_df1['Active Risk']**2)
    Setup_df1['Geometric Total'] = np.exp(np.log(1+ Setup_df1.iloc[:,9])-(np.log(1+((Setup_df1.iloc[:,10]**2)/((1+Setup_df1.iloc[:,9])**2))))/2)-1

    return Setup_df1


# Passive Correlation Table (Excel formula is yielding wrong results)

# In[5]:


def Setup_df2(file):
    
    X = Setup_df1(file)
    Y = CMA_df2(file)

    Setup_df2 = pd.DataFrame(columns = X.index.T, index = X.index)

    for i in X.index:
        for j in X.index.T:
            Setup_df2.loc[i,j] = Y.loc[i,j]
    
    return Setup_df2


# Active Correlation dataframe

# In[6]:


def Setup_df3(file):
    
    X = Setup_df1(file)

    Setup_df3 = pd.DataFrame(columns = X.index.T, index = X.index)

    for i in X.index:
        for j in X.index.T:
            if i == j:
                Setup_df3.loc[i,j] = 1
            else:
                Setup_df3.loc[i,j] = 0
    
    return Setup_df3
            


# Passive Covaiance Matrix

# In[7]:


def Setup_df5(file):
    
    X = Setup_df1(file)
    Y = Setup_df2(file)
    
    Passive_vol = X.iloc[:,7]
    arr = Passive_vol.to_numpy()
    arr = arr.reshape(-1,1)
    arr_transp = arr.T
    mat_mult = np.dot(arr, arr_transp)
    df_mat_mult = pd.DataFrame(mat_mult, columns = X.index.T, index = X.index)

    Setup_df5 = pd.DataFrame(columns = X.index.T, index = X.index)

    for i in X.index:
        for j in X.index.T:
            Setup_df5.loc[i,j] = df_mat_mult.loc[i,j] * Y.loc[i,j]

    return Setup_df5


# Active Covariance Matrix

# In[8]:


def Setup_df6(file):
    
    X = Setup_df1(file)
    Y = Setup_df3(file)
    
    Total_vol = X.iloc[:,10]
    arr1 = Total_vol.to_numpy()
    arr1 = arr1.reshape(-1,1)
    arr1_transp = arr1.T
    mat_mult1 = np.dot(arr1, arr1_transp)
    df_mat_mult1 = pd.DataFrame(mat_mult1, columns = X.index.T, index = X.index)

    Setup_df6 = pd.DataFrame(columns = X.index.T, index = X.index)

    for i in X.index:
        for j in X.index.T:
            Setup_df6.loc[i,j] = df_mat_mult1.loc[i,j] * Y.loc[i,j]

    return Setup_df6


# Total Covariance Matrix

# In[9]:


def Setup_df7(file):
    
    X = Setup_df5(file)
    Y = Setup_df6(file)

    Setup_df7 = X + Y
    return Setup_df7


# Allocation Input Table

# In[10]:


#Function for reading Input allocation table

def Allocation_df1(file):
    
    Allocation_df1 = pd.read_excel(file, sheet_name=0, usecols="B:W", header=3, index_col=0, nrows=28)
    
    return Allocation_df1


# Output Table (Risk-Return Calculation)

# In[11]:


def Risk_ret_output(file):
    
    '''
    Returns the Risk Return Calculation Output Table.

    Parameters:
        file:The path of excel file where raw data and allocation inputs are saved.

    Returns:
      Risk_ret_output(file): Risk Return Calculation Output Table   
    '''
    
    A = Setup_df1(file)
    B = Allocation_df1(file)
    C = Setup_df5(file)
    D = Setup_df6(file)
    E = Setup_df7(file)


    Risk_ret_output = pd.read_excel(file, sheet_name=0, usecols="B:W", header=33, index_col=0, nrows=12)
    Risk_ret_output = Risk_ret_output.T.applymap(lambda x: 0)

    for i in range(0,21):
        Risk_ret_output.iloc[i,0]  = A['Arithmetic Total'].to_numpy().reshape(-1,1).T.dot(B.iloc[:,i].to_numpy().reshape(-1,1))[0]
        Risk_ret_output.iloc[i,2]  = (multi_dot([B.iloc[:,i].to_numpy().reshape(-1,1).T, E.to_numpy(), B.iloc[:,i].to_numpy().reshape(-1,1)])[0])**(0.5)
        Risk_ret_output.iloc[i,3]  = Risk_ret_output.iloc[i,0] / Risk_ret_output.iloc[i,2]
        Risk_ret_output.iloc[i,5]  = A['Active Return'].to_numpy().reshape(-1,1).T.dot(B.iloc[:,i].to_numpy().reshape(-1,1))[0]
        Risk_ret_output.iloc[i,6]  = (multi_dot([B.iloc[:,i].to_numpy().reshape(-1,1).T, D.to_numpy(), B.iloc[:,i].to_numpy().reshape(-1,1)])[0])**(0.5)
        Risk_ret_output.iloc[i,7]  = A['Fee'].to_numpy().reshape(-1,1).T.dot(B.iloc[:,i].to_numpy().reshape(-1,1))[0]
        Risk_ret_output.iloc[i,4]  = Risk_ret_output.iloc[i,5] - Risk_ret_output.iloc[i,7]
        Risk_ret_output.iloc[i,1]  = np.exp(np.log(1+ Risk_ret_output.iloc[i,0])-(np.log(1+((Risk_ret_output.iloc[i,2]**2)/((1+Risk_ret_output.iloc[i,0])**2))))/2)-1-Risk_ret_output.iloc[i,7]
        Risk_ret_output.iloc[i,8]  = A['Arithmetic Passive'].to_numpy().reshape(-1,1).T.dot(B.iloc[:,i].to_numpy().reshape(-1,1))[0]
        Risk_ret_output.iloc[i,10] = (multi_dot([B.iloc[:,i].to_numpy().reshape(-1,1).T, C.to_numpy(), B.iloc[:,i].to_numpy().reshape(-1,1)])[0])**(0.5)
        Risk_ret_output.iloc[i,9]  = np.exp(np.log(1+ Risk_ret_output.iloc[i,8])-(np.log(1+((Risk_ret_output.iloc[i,10]**2)/((1+Risk_ret_output.iloc[i,8])**2))))/2)-1
        Risk_ret_output.iloc[i,11] = Risk_ret_output.iloc[i,9] - A['Geometric Passive'].to_numpy().reshape(-1,1).T.dot(B.iloc[:,i].to_numpy().reshape(-1,1))[0]


    Risk_ret_output = Risk_ret_output.fillna(0).T.style.format('{:,.2%}')
    return Risk_ret_output

Risk_ret_output(file)


# Generating Climate Change Stress Tests Table

# Function for reading Input allocation table

# In[12]:


#Function for reading Input allocation table

file =(r'C:\Users\Gaurav\Desktop\Python_Assignment.xlsx')

def Allocation_df1(file):
    
    Allocation_df1 = pd.read_excel(file, sheet_name=0, usecols="B:W", header=3, index_col=0, nrows=28)
    
    return Allocation_df1


# Input dataframes for Scenario sheet

# Function for building initial dataframe (CC_Scen_df1) used for building intermediatce Climate Change Stress Tests Expected Cumulative Returns dataframe (Setup_df4)

# In[13]:


#Function for building initial dataframe (CC_Scen_df1) used for building intermediatce Climate Change Stress Tests Expected Cumulative Returns dataframe (Setup_df4)

def CC_Scen_df1(file):
    
    CC_Scen_df1 = pd.read_excel(file, sheet_name=3, usecols="A:C,E:N", header=3, index_col=2, nrows=232)
    CC_Scen_df1 = CC_Scen_df1.iloc[1:]
    CC_Scen_df1['1y Cum Ret'] = CC_Scen_df1.iloc[:,2]
    CC_Scen_df1['3y Cum Ret'] = (CC_Scen_df1.iloc[:,2:5].apply(lambda x: x+1).prod(axis=1))**(1/3)-1
    CC_Scen_df1['5y Cum Ret'] = (CC_Scen_df1.iloc[:,2:7].apply(lambda x: x+1).prod(axis=1))**(1/5)-1
    CC_Scen_df1['10y Cum Ret'] = (CC_Scen_df1.iloc[:,2:12].apply(lambda x: x+1).prod(axis=1))**(1/10)-1
    
    return CC_Scen_df1


# Function for building Climate Change Stress Tests Expected Cumulative Returns dataframe (Setup_df4)

# In[14]:


#Function for building Climate Change Stress Tests Expected Cumulative Returns dataframe (Setup_df4)

def Setup_df4(file):
    
    A = CC_Scen_df1(file)

    Setup_df4 = pd.read_excel(file, sheet_name=1, usecols="BY", header=1, nrows=28)

    Setup_df4.columns = ["Portfolio Name"]
    Setup_df4["Scenario 1"] = "Transition (2°C)"
    Setup_df4["Scenario 2"] = "Low Mitigation (4°C)"
    Setup_df4["Index 1"] = Setup_df4["Scenario 1"].str.cat(Setup_df4["Portfolio Name"])
    Setup_df4["Index 2"] = Setup_df4["Scenario 2"].str.cat(Setup_df4["Portfolio Name"])

    Setup_df4['Transition (2°C) 1y Cum Ret'] = Setup_df4["Index 1"].map(A['1y Cum Ret'])
    Setup_df4['Transition (2°C) 3y Cum Ret'] = Setup_df4["Index 1"].map(A['3y Cum Ret'])
    Setup_df4['Transition (2°C) 5y Cum Ret'] = Setup_df4["Index 1"].map(A['5y Cum Ret'])
    Setup_df4['Transition (2°C) 10y Cum Ret'] = Setup_df4["Index 1"].map(A['10y Cum Ret'])

    Setup_df4['Low Mitigation (4°C) 1y Cum Ret'] = Setup_df4["Index 2"].map(A['1y Cum Ret'])
    Setup_df4['Low Mitigation (4°C) 3y Cum Ret'] = Setup_df4["Index 2"].map(A['3y Cum Ret'])
    Setup_df4['Low Mitigation (4°C) 5y Cum Ret'] = Setup_df4["Index 2"].map(A['5y Cum Ret'])
    Setup_df4['Low Mitigation (4°C) 10y Cum Ret'] = Setup_df4["Index 2"].map(A['10y Cum Ret'])

    Setup_df4.set_index(Setup_df4.iloc[:,0].name, inplace= True)
    Setup_df4 = Setup_df4.iloc[:,4:]
    
    return Setup_df4


# Function for producing Climate Change Stress Tests Output Table

# In[15]:


# Function for producing Climate Change Stress Tests Output Table


def CC_output(file):
    
    '''
    Returns the Climate Change Stress Tests Output Table.

    Parameters:
        file:The path of excel file where raw data and allocation inputs are saved.

    Returns:
      CC_output(file): Climate Change Stress Tests Output Table   
    '''
    
    #Calling function CC_Scen_df1 used for building initial dataframe (CC_Scen_df1) further used in building intermediatce Climate Change Stress Tests Expected Cumulative Returns dataframe (Setup_df4)
    
    CC_Scen_df1(file)
    
    #Calling function Setup_df4 used for for building Climate Change Stress Tests Expected Cumulative Returns dataframe (Setup_df4)
    
    X = Setup_df4(file)
    
    
    #Calling function Allocation_df1 from CC_functions.py file used for reading Input allocation table
    
    
    Y = Allocation_df1(file)

    
    #Code for producing Climate change stress tests output table
    
    CC_output = pd.read_excel(file, sheet_name=0, usecols="B:W", header=48, index_col=0, nrows=12)
    CC_output = CC_output.T

    for i in range(0,21):
        CC_output.iloc[i,1]  = X['Transition (2°C) 1y Cum Ret'].to_numpy().reshape(-1,1).T.dot(Y.iloc[:,i].to_numpy().reshape(-1,1))[0]
        CC_output.iloc[i,2]  = X['Low Mitigation (4°C) 1y Cum Ret'].to_numpy().reshape(-1,1).T.dot(Y.iloc[:,i].to_numpy().reshape(-1,1))[0]
        CC_output.iloc[i,4]  = X['Transition (2°C) 3y Cum Ret'].to_numpy().reshape(-1,1).T.dot(Y.iloc[:,i].to_numpy().reshape(-1,1))[0]
        CC_output.iloc[i,5]  = X['Low Mitigation (4°C) 3y Cum Ret'].to_numpy().reshape(-1,1).T.dot(Y.iloc[:,i].to_numpy().reshape(-1,1))[0]
        CC_output.iloc[i,7]  = X['Transition (2°C) 5y Cum Ret'].to_numpy().reshape(-1,1).T.dot(Y.iloc[:,i].to_numpy().reshape(-1,1))[0]
        CC_output.iloc[i,8]  = X['Low Mitigation (4°C) 5y Cum Ret'].to_numpy().reshape(-1,1).T.dot(Y.iloc[:,i].to_numpy().reshape(-1,1))[0]
        CC_output.iloc[i,10] = X['Transition (2°C) 10y Cum Ret'].to_numpy().reshape(-1,1).T.dot(Y.iloc[:,i].to_numpy().reshape(-1,1))[0]
        CC_output.iloc[i,11] = X['Low Mitigation (4°C) 10y Cum Ret'].to_numpy().reshape(-1,1).T.dot(Y.iloc[:,i].to_numpy().reshape(-1,1))[0]

    pd.options.display.float_format = '{:.2%}'.format    
    CC_output = CC_output.T.fillna('')
    
    return CC_output

CC_output(file)


# Dashboard creation

# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:



#from CC_functions import CC_Scen_df1, Setup_df4, Allocation_df1


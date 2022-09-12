# -*- coding: utf-8 -*-
"""
Created on Thu Jul 21 19:53:52 2022

@author: ba331c
"""

from statsmodels.tsa.statespace.varmax import VARMAX
from statsmodels.tsa.api import VAR
from statsmodels.tsa.stattools import grangercausalitytests, adfuller
from collections import Counter

from sklearn.metrics import mean_squared_error
import math 
from statistics import mean

import matplotlib.pyplot as plt
import pandas as pd
from pandas import Series, ExcelWriter

import numpy as np


counter = 8
with ExcelWriter('Rezultati_Proizvodi.xlsx') as writer:
    for i in range(10,11 ):
        file_name = "C:\\Users\\ba331c\\Desktop\\Lazo Phd\\S-Proizvodi.xlsx"
        sheet =  i
        file_info = pd.ExcelFile(file_name)
        data = pd.read_excel(io=file_name, sheet_name=sheet)
        sheet_names = file_info.sheet_names
        col = 1
        country_lag= {}
        ad_fuller_results=[]
        
        for col in range(1,data.shape[1]):
            #proveruvame stationarity na nasite variabli
            ad_fuller_result_1 = adfuller(data.iloc[:,col].diff()[1:])
            p_value_stationarity=ad_fuller_result_1[1]
            country = data.iloc[:,col].name
            ad_fuller_results.append(ad_fuller_result_1)
            
            p_values={}
            #proveruvame kauzalnost
            print(sheet_names[i],country)
            granger_2 = grangercausalitytests(data.iloc[:,np.r_[1,col]], 6)
            
            lag = 1
            granger_causality_results=[]
            for lag in range(1,6):
                #granger_causality_results.append(granger_2[lag][0]['ssr_ftest'][1])
                if granger_2[lag][0]['ssr_ftest'][1]<0.05:
                    p_values[lag] = granger_2[lag][0]['ssr_ftest'][1]
            country_lag[country] = p_values
        
        
        
        
        useful_country=['Macedonia']
        for u in country_lag:
            if len(country_lag[u])>0:
                useful_country.append(u)
            
        subdata=data[useful_country]
        train_test_split = int(round(data.shape[0]*0.1,0))
        train_df=subdata[:-train_test_split]
        test_df=subdata[-train_test_split:]
        
    
        #proveruvame koj time lag dava najgolema korelacija
        model = VAR(train_df.diff()[1:])
        sorted_order=model.select_order(maxlags=6)
        
        df_lag= pd.DataFrame(sorted_order.summary())
        
        all_lags=[]
        for l in sorted_order.selected_orders:
            all_lags.append(sorted_order.selected_orders[l])
            cnt = Counter()
            for number in all_lags:
                 cnt[number] += 1
            
        lag_to_use=max(cnt, key=cnt.get)       
        
        if lag_to_use==0:
            lag_to_use=1
        else:
            lag_to_use
        
        var_model = VARMAX(train_df, order=(lag_to_use,0))
        fitted_model = var_model.fit(disp=False)
        
        
        
        n_forecast = train_test_split + 12
        predict = fitted_model.get_prediction(start=len(train_df),\
                                              end=len(train_df) + n_forecast-1)
        predictions=predict.predicted_mean
        predictions.rename(columns = {'Macedonia' : 'Macedonia_predicted'},\
                           inplace = True)
          
        test_vs_pred=pd.concat([test_df['Macedonia'],\
                                predictions['Macedonia_predicted']],axis=1)
        
    
        df_adf=pd.DataFrame(ad_fuller_results)  
        df_adf.to_excel(writer,\
                        sheet_name = sheet_names[i] + ' rezultati' ,startcol=1,startrow=1)
        #df_granger=pd.DataFrame(granger_causality_results)
        #df_granger.to_excel(writer,\
        #        sheet_name = sheet_names[i] + ' rezultati' ,startcol=1,startrow=10)
        df_lag.to_excel(writer,\
                        sheet_name = sheet_names[i] + ' rezultati' ,startcol=1,startrow=20)
        data.to_excel(writer, \
                      sheet_name=sheet_names[i] + ' rezultati' ,startcol=13)
        test_vs_pred.to_excel(writer,\
                              sheet_name=sheet_names[i] + ' rezultati',startcol=30)
    
    writer.save()



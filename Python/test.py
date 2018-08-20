# -*- coding: utf-8 -*-
"""
Created on Fri Aug 10 10:22:54 2018

@author: Najib
"""

## From SQL to DataFrame Pandas
import pandas as pd
import pyodbc

sql_conn = pyodbc.connect('DRIVER={SQL Server};SERVER=52.187.121.85;PORT=1433;DATABASE=LeadIntelligenceTest;uid=sa;pwd=1Qaz2wsx3edc;Trusted_Connection=NO')
query = "SELECT * FROM [TB_A_MAPPING]"
df = pd.read_sql(query, sql_conn)

print(df)
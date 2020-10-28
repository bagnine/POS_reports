import pandas as pd
import numpy as np
import matplotlib.pyplot as plt 
import seaborn as sns
import xlsxwriter

summary = pd.read_csv('Sales_Summary_restaurant_2020-10-06_11-00_2020-10-25_11-00.csv')
pmix = pd.read_csv('Product_Mix_restaurant_2020-10-06_11-00_2020-10-25_11-00.csv')
hourly = pd.read_csv('Hourly_Sales_restaurant_2020-10-06_11-00_2020-10-25_11-00.csv')
emp_prof = pd.read_csv('Employee_Profit_restaurant_2020-10-06_2020-10-25.csv')
labor = pd.read_csv('Labor_restaurant_2020-10-06_2020-10-25.csv')

def category(df, category):
    a = df.iloc[np.where(df['Product Category'] == category)]
    return a[['Name', 'Product Subcategory', 'Qty', 'Total Sales']]

summary['Time From'] = summary['Time From'].apply(lambda x: str(x))
summary['Time From'] = summary['Time From'].apply(lambda x: x[0:8])

    #reformatting the date column to keep only relevant information

df1 = summary[['Time From','Gross Sales', 'Net Sales']]

df2 = labor[['Unnamed: 0', 'Actual Hours', '# Transactions', 'Wage', 'Sales']]

emp_prof.drop(columns = 'Employee', inplace = True)

    #redacting the employee names for anonymity

df3 = emp_prof[['Labor', 'Hours', 'Sales', 'Cost', 'Profit']]

df4 = hourly[hourly['# Transactions']!= 0]

liquor = category(pmix, 'Liquor')
beer = category(pmix, 'Beer')
food = category(pmix, 'Food')

gross = df1.iloc[-1][2]
net = df1.iloc[-1][2]
wages = df2.iloc[-1][3]
hours = df3.iloc[-1][1]
liquorsales = liquor['Total Sales'].sum()
beersales = beer['Total Sales'].sum()
foodsales = food['Total Sales'].sum()

f = {'Gross Sales': gross, 'Net Sales': net, 'Total Wages': wages, 
    'Employee Hours': hours, 'Liquor Sales': liquorsales, 
    'Beer Sales': beersales, 'Food Sales': foodsales}
    
frontset = pd.Series(f)

def export_sheets():
    writer = pd.ExcelWriter('restaurant.xlsx', engine='xlsxwriter')

    frontset.to_excel(writer, sheet_name= 'Totals')
    df1.to_excel(writer, sheet_name='Summary')
    df2.to_excel(writer, sheet_name='Labor')
    df3.to_excel(writer, sheet_name='Hours')
    df4.to_excel(writer, sheet_name='Hourly Sales')
    liquor.to_excel(writer, sheet_name= 'Liquor')
    beer.to_excel(writer, sheet_name= 'Beer')
    food.to_excel(writer, sheet_name= 'Food')

    return writer.save()

export_sheets()
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# Load the dataset

df = pd.read_csv('SuperStore_Sales_Dataset.csv', encoding='utf-8')

# Initial exploration

print("Dataset Shape:", df.shape)
print("\nFirst 5 rows:")
print(df.head())
print("\nDataset Info:")
print(df.info())
print("\nMissing Values:")
print(df.isnull().sum())
print("\nBasic Statistics:")
print(df.describe())


# Data Cleaning and preprocessing

df['Order Date'] = pd.to_datetime(df['Order Date'], format='%d-%m-%Y')
df['Ship Date'] = pd.to_datetime(df['Ship Date'], format='%d-%m-%Y')

# Extract time features
df['Order Year'] = df['Order Date'].dt.year
df['Order Month'] = df['Order Date'].dt.month
df['Order Quarter'] = df['Order Date'].dt.quarter
df['Order Day'] = df['Order Date'].dt.day
df['Order Weekday'] = df['Order Date'].dt.day_name()
df['Order Month Name'] = df['Order Date'].dt.strftime('%B')



# Calculate shipping duration
df['Shipping Days'] = (df['Ship Date'] - df['Order Date']).dt.days



# Handle missing values in Returns column (assuming #N/A means no return)
df['Returns'] = df['Returns'].replace('#N/A', 'No')
df['Returns'] = df['Returns'].fillna('No')


df['Profit'] = pd.to_numeric(df['Profit'], errors='coerce')
df['Return Flag'] = df['Returns'].apply(lambda x: 'Yes' if x != 'No' else 'No')

# Calculate profit margin
df['Profit Margin'] = (df['Profit'] / df['Sales'] * 100).r
df['Sales Category'] = pd.cut(df['Sales'], bins=[0, 100, 500, 1000, float('inf')],labels=['Low (<$100)', 'Medium ($100-$500)', 'High ($500-$1000)', 'Very High (>$1000)'])



# Calculation

total_sales = df['Sales'].sum()
total_profit = df['Profit'].sum()
total_quantity = df['Quantity'].sum()
avg_profit_margin = df['Profit Margin'].mean()
avg_shipping_days = df['Shipping Days'].mean()
return_rate = (df['Return Flag'] == 'Yes').sum() / len(df) * 100

print(f"Total Sales: ${total_sales:,.2f}")
print(f"Total Profit: ${total_profit:,.2f}")
print(f"Total Quantity Sold: {total_quantity:,}")
print(f"Average Profit Margin: {avg_profit_margin:.2f}%")
print(f"Average Shipping Days: {avg_shipping_days:.2f}")
print(f"Return Rate: {return_rate:.2f}%")

# Top performing categories
category_performance = df.groupby('Category').agg({
    'Sales': 'sum',
    'Profit': 'sum',
    'Quantity': 'sum'
}).sort_values('Sales', ascending=False)

print("\nCategory Performance:")
if(print(category_performance)):
    print("Done")

monthly_sales = df.groupby(['Order Year', 'Order Month', 'Order Month Name']).agg({
    'Sales': 'sum',
    'Profit': 'sum',
    'Quantity': 'sum'
}).reset_index()

category_sales = df.groupby(['Category', 'Sub-Category']).agg({
    'Sales': 'sum',
    'Profit': 'sum',
    'Quantity': 'sum',
    'Row ID': 'count'
}).rename(columns={'Row ID': 'Transaction Count'}).reset_index()

regional_performance = df.groupby(['Region', 'State']).agg({
    'Sales': 'sum',
    'Profit': 'sum',
    'Customer ID': 'nunique'
}).rename(columns={'Customer ID': 'Unique Customers'}).reset_index()

customer_segmentation = df.groupby(['Segment', 'Customer ID', 'Customer Name']).agg({
    'Sales': 'sum',
    'Profit': 'sum',
    'Order ID': 'nunique'
}).rename(columns={'Order ID': 'Order Count'}).reset_index()

shipping_analysis = df.groupby(['Ship Mode', 'Region']).agg({
    'Sales': 'sum',
    'Shipping Days': 'mean',
    'Order ID': 'count'
}).rename(columns={'Order ID': 'Shipment Count'}).reset_index()

payment_analysis = df.groupby('Payment Mode').agg({
    'Sales': 'sum',
    'Profit': 'sum',
    'Order ID': 'count'
}).rename(columns={'Order ID': 'Transaction Count'}).reset_index()

product_performance = df.groupby(['Product ID', 'Product Name', 'Category', 'Sub-Category']).agg({
    'Sales': 'sum',
    'Profit': 'sum',
    'Quantity': 'sum',
    'Profit Margin': 'mean'
}).reset_index()

returns_analysis = df[df['Return Flag'] == 'Yes'].groupby(['Category', 'Sub-Category']).agg({
    'Sales': 'sum',
    'Order ID': 'count'
}).rename(columns={'Order ID': 'Return Count'}).reset_index()





with pd.ExcelWriter('SuperStore_Analysis.xlsx') as writer:
    df.to_excel(writer, sheet_name='Main Data', index=False)
    monthly_sales.to_excel(writer, sheet_name='Monthly Sales', index=False)
    category_sales.to_excel(writer, sheet_name='Category Sales', index=False)
    regional_performance.to_excel(writer, sheet_name='Regional Performance', index=False)
    customer_segmentation.to_excel(writer, sheet_name='Customer Segmentation', index=False)
    shipping_analysis.to_excel(writer, sheet_name='Shipping Analysis', index=False)
    payment_analysis.to_excel(writer, sheet_name='Payment Analysis', index=False)
    product_performance.to_excel(writer, sheet_name='Product Performance', index=False)
    returns_analysis.to_excel(writer, sheet_name='Returns Analysis', index=False)

print("All analysis data saved to 'SuperStore_Analysis.xlsx'")

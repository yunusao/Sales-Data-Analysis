import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl.drawing.image import Image

def generate_sales_data():
    np.random.seed(0)
    dates = pd.date_range(start='2023-01-01', end='2023-12-31', freq='D')
    products = ['Product A', 'Product B', 'Product C', 'Product D']
    regions = ['North', 'South', 'East', 'West']

    data = {
        'Date': np.random.choice(dates, size=1000),
        'Product': np.random.choice(products, size=1000),
        'Quantity Sold': np.random.randint(1, 20, size=1000),
        'Price per Unit': np.random.uniform(10, 100, size=1000).round(2),
        'Region': np.random.choice(regions, size=1000)
    }

    sales_data = pd.DataFrame(data)
    sales_data['Total Sales'] = sales_data['Quantity Sold'] * sales_data['Price per Unit']
    sales_data['Month'] = sales_data['Date'].dt.to_period('M')
    sales_data['Month'] = sales_data['Month'].astype(str)
    return sales_data

def perform_data_analysis(sales_data):
    total_sales = sales_data['Total Sales'].sum()
    average_sales = sales_data['Total Sales'].mean()
    total_quantity_sold = sales_data['Quantity Sold'].sum()
    average_quantity_sold = sales_data['Quantity Sold'].mean()

    sales_by_product = sales_data.groupby('Product')['Total Sales'].sum().sort_values(ascending=False)
    average_sales_by_product = sales_data.groupby('Product')['Total Sales'].mean().sort_values(ascending=False)

    sales_by_region = sales_data.groupby('Region')['Total Sales'].sum().sort_values(ascending=False)
    average_sales_by_region = sales_data.groupby('Region')['Total Sales'].mean().sort_values(ascending=False)

    sales_by_month = sales_data.groupby('Month')['Total Sales'].sum().reset_index()

    return {
        'total_sales': total_sales,
        'average_sales': average_sales,
        'total_quantity_sold': total_quantity_sold,
        'average_quantity_sold': average_quantity_sold,
        'sales_by_product': sales_by_product,
        'average_sales_by_product': average_sales_by_product,
        'sales_by_region': sales_by_region,
        'average_sales_by_region': average_sales_by_region,
        'sales_by_month': sales_by_month
    }

def create_visualizations(sales_data, analysis_results):
    sales_by_product = analysis_results['sales_by_product']
    sales_by_region = analysis_results['sales_by_region']
    sales_by_month = analysis_results['sales_by_month']

    plt.figure(figsize=(10, 6))
    sales_by_product.plot(kind='bar', color='skyblue')
    plt.title('Total Sales by Product')
    plt.xlabel('Product')
    plt.ylabel('Total Sales')
    plt.xticks(rotation=45)
    plt.savefig('total_sales_by_product.png')
    plt.close()

    plt.figure(figsize=(10, 6))
    sns.barplot(x=sales_by_region.index, y=sales_by_region.values, palette='viridis')
    plt.title('Total Sales by Region')
    plt.xlabel('Region')
    plt.ylabel('Total Sales')
    plt.xticks(rotation=45)
    plt.savefig('total_sales_by_region.png')
    plt.close()

    fig = px.line(sales_by_month, x='Month', y='Total Sales', title='Total Sales Over Time')
    fig.write_image('total_sales_over_time.png')

    plt.figure(figsize=(10, 6))
    sns.boxplot(x='Product', y='Total Sales', data=sales_data, palette='Set2')
    plt.title('Average Sales per Transaction by Product')
    plt.xlabel('Product')
    plt.ylabel('Total Sales')
    plt.xticks(rotation=45)
    plt.savefig('average_sales_per_transaction.png')
    plt.close()

def export_to_excel(sales_data, analysis_results):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sales Data Analysis"

    ws.append(["Key Metrics"])
    ws.append(["Total Sales", analysis_results['total_sales']])
    ws.append(["Average Sales per Transaction", analysis_results['average_sales']])
    ws.append(["Total Quantity Sold", analysis_results['total_quantity_sold']])
    ws.append(["Average Quantity Sold per Transaction", analysis_results['average_quantity_sold']])
    ws.append([])

    ws.append(["Sales by Product"])
    for r in dataframe_to_rows(analysis_results['sales_by_product'].reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])

    ws.append(["Average Sales by Product"])
    for r in dataframe_to_rows(analysis_results['average_sales_by_product'].reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])

    ws.append(["Sales by Region"])
    for r in dataframe_to_rows(analysis_results['sales_by_region'].reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])

    ws.append(["Average Sales by Region"])
    for r in dataframe_to_rows(analysis_results['average_sales_by_region'].reset_index(), index=False, header=True):
        ws.append(r)
    ws.append([])

    ws.append(["Sales Over Time"])
    for r in dataframe_to_rows(analysis_results['sales_by_month'], index=False, header=True):
        ws.append(r)

    wb.save("Sales_Data_Analysis.xlsx")

    wb = openpyxl.load_workbook("Sales_Data_Analysis.xlsx")
    ws = wb.active

    img1 = Image('total_sales_by_product.png')
    img2 = Image('total_sales_by_region.png')
    img3 = Image('total_sales_over_time.png')
    img4 = Image('average_sales_per_transaction.png')

    ws.add_image(img1, 'G2')
    ws.add_image(img2, 'G20')
    ws.add_image(img3, 'G38')
    ws.add_image(img4, 'G56')

    wb.save("Sales_Data_Analysis_with_Visualizations.xlsx")

def main():
    sales_data = generate_sales_data()
    analysis_results = perform_data_analysis(sales_data)
    create_visualizations(sales_data, analysis_results)
    export_to_excel(sales_data, analysis_results)

if __name__ == "__main__":
    main()

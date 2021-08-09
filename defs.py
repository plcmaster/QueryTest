#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Functions Definitions
Name: defs.py

   Functions definition for get the Shopify e-commerse
   statistics trough REST API
   
"""

__author__ = "Pedro Campana"
__date__ = "2021-07-31"
__version__ = "1.0.0"
__maintainer__ = "Pedro Campana"
__email__ = "pcampana.pc@gmail.com"
__status__ = "Development"


import requests
import json
import pandas as pd
from requests.models import LocationParseError
import xlwings as xw
import openpyxl
import os
import win32com.client as win32
from pathlib import Path, PureWindowsPath
import time
import datetime


def ordersummary(ordersData):

        data = ordersData 

        ordersTotal = data.groupby('order_number')['line_items.price'].sum().reset_index()
        shippingCost = data.groupby('order_number')['discounted_price'].mean().reset_index()

        ordersTotal.set_index('order_number')
        shippingCost.set_index('order_number')

        orderTotalSummary = ordersTotal.merge(shippingCost, on='order_number')

        orderTotalSummary =orderTotalSummary.rename(columns={'line_items.price':'Order_Price','discounted_price':'Shipping_Cost'})
        return orderTotalSummary

    

def get_orders(url,fromdate, todate):
  
        if not todate:
                now = datetime.date.today()
                enddate=now.strftime("%Y-%m-%dT%H:%M:%S-00:00")
        else:
                enddate = todate
        
        endpoint = "orders.json?status=any&limit=250&created_at_min=" + fromdate + "&" + "created_at_max=" + enddate + "fields=order_number,created_at,status,line_items,total_price,shipping_address,total_shipping_price_set"
        r = requests.get(url + endpoint)
        print(url+endpoint)
       
        orders = r.json()
        
        data = pd.json_normalize(data=orders['orders'], meta=['order_number','total_price',
        'shipping_address','created_at','shipping_lines'],record_path=['line_items'],record_prefix='line_items.', max_level=1)
       
        data_1 = data.drop('shipping_address',1).assign(**pd.DataFrame(data.shipping_address.values.tolist()))
        
        data_2 = data_1.drop('shipping_lines',1).assign(**pd.DataFrame(data_1.shipping_lines.values.tolist()))
       
        data_2['line_items.price'] = pd.to_numeric(data_2['line_items.price'])
        data_2['total_price'] = pd.to_numeric(data_2['total_price'])
        data_2['discounted_price'] = pd.to_numeric(data_2['discounted_price'])

        # Columnas removidas del dataframe original, no necesarias para el análisis
        #**************************************************************************
        data_2 = data_2.drop([
                "line_items.admin_graphql_api_id",  
                "line_items.fulfillable_quantity",
                "line_items.fulfillment_service",
                "line_items.grams",
                "line_items.variant_inventory_management",
                "line_items.tax_lines",
                "line_items.price_set.shop_money",
                "line_items.price_set.presentment_money",
                "line_items.total_discount_set.presentment_money",
                "line_items.title",
                "line_items.origin_location.country_code",
                "line_items.origin_location.province_code",
                "line_items.origin_location.name",
                "line_items.origin_location.address1",
                "line_items.origin_location.address2",
                "line_items.origin_location.city",
                "line_items.origin_location.zip",
                "discounted_price_set",
                "price",
                "price_set",
                "tax_lines",
                "discount_allocations",
                "requested_fulfillment_service_id",
                "id",
                "carrier_identifier",
                "delivery_category",
                "line_items.properties",
                "line_items.duties",
                "line_items.discount_allocations",
                "line_items.tip_payment_gateway",
                "line_items.tip_payment_method",
                "line_items.tip.payment_method",
                "line_items.tip.payment_gateway",
                # "line_items.destination_location.id",
                "line_items.total_discount_set.shop_money",
                # "line_items.destination_location.country_code",
                # "line_items.destination_location.province_code",
                # "line_items.destination_location.name",
                # "line_items.destination_location.address1",
                # "line_items.destination_location.address2",
                # "line_items.destination_location.city",
                # "line_items.destination_location.zip",
                # "line_items.tip_payment_gateway",
                # "line_items.tip_payment_method",
                # "line_items.tip.payment_method",
                # "line_items.tip.payment_gateway"

                
                ], axis=1
                )

         #*************************************************************************

        columnsToMove = ['order_number','line_items.sku','line_items.name','line_items.price','total_price','discounted_price'] 
        data_2 = data_2[ columnsToMove + [ col for col in data_2.columns if col not in columnsToMove ] ]

        return data_2


def get_inventory(url):

        endpoint = "products.json?&limit=250"
        r = requests.get(url + endpoint)

        products = r.json()

        data = pd.json_normalize(data=products['products'], meta=['title','product_type','created_at',
        'updated_at','published_at'],record_path=['variants'],record_prefix='variants.', max_level=1)

        # Columnas removidas del dataframe original, no necesarias para el análisis
        #************************
        data = data.drop([
                "variants.id",
                "variants.product_id",
                "variants.admin_graphql_api_id",  
                "variants.inventory_policy",
                "variants.compare_at_price",
                "variants.fulfillment_service",
                "variants.inventory_management",
                "variants.option1",
                "variants.option2",
                "variants.option3",
                "variants.taxable",
                "variants.barcode",
                "variants.grams",
                "variants.image_id",
                "variants.weight",
                "variants.weight_unit",
                "variants.created_at",
                "variants.updated_at",
                "created_at",
                "updated_at",
                "variants.position",
                "variants.old_inventory_quantity"
                ], axis=1
                )
        #***************************
        data['variants.price'] = pd.to_numeric(data['variants.price'])
       
        columnsToMove = ['variants.inventory_item_id','variants.sku','title'] 
        data = data[ columnsToMove + [ col for col in data.columns if col not in columnsToMove ] ]

        # Se calcula la valorización de stock por producto y se agrega nueva columna
        data['Stock-Value'] = data['variants.price'] * data['variants.inventory_quantity']
        data.rename(columns={'variants.sku':'sku'}, inplace=True)


        inventoryItemIDList = data['variants.inventory_item_id'].values.tolist()
        inventoryItemIDListstr = str(inventoryItemIDList)
        inventoryItemIDListstr = inventoryItemIDListstr.replace('[','')
        inventoryItemIDListstr = inventoryItemIDListstr.replace(']','')

        return data, inventoryItemIDListstr


def get_inventoryItem(idlist,url):

        endpoint = "inventory_items.json?ids=" + idlist +"&limit=250"
        r = requests.get( url + endpoint)
        
        inventory = r.json()

        inventoryItem_df = pd.json_normalize(data=inventory['inventory_items'])

        inventoryItem_df = inventoryItem_df.drop(["created_at",
                                "updated_at",
                                "requires_shipping",
                                
                                "country_code_of_origin",
                                "province_code_of_origin",
                                "harmonized_system_code",
                                "tracked",
                                "country_harmonized_system_codes",
                                "admin_graphql_api_id"
                                ], axis=1)
        inventoryItem_df['cost'] = pd.to_numeric(inventoryItem_df['cost'])

        return inventoryItem_df


def dfToexcel(excelfile,sheet,df):

        writer = pd.ExcelWriter(excelfile, engine='xlsxwriter')
                
        df.to_excel(writer,sheet_name=sheet)

        workbook = writer.book
        worksheet = writer.sheets[sheet]

        format1 = workbook.add_format({'num_format':'#######################'})
        format2 = workbook.add_format({'num_format':'$###,###'})

        # worksheet.set_column('H:H',20,format1)
        # worksheet.set_column('L:L',20,format1)
        # worksheet.set_column('R:R',20,format1)
        # worksheet.set_column('W:W',20,format1)
        # worksheet.set_column('E:E',20,format2)
        # worksheet.set_column('F:F',20,format2)
        # worksheet.set_column('G:G',20,format2)
        
        writer.save()
        
        

  

def setcolumnwidth(file,sheet):

                   
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(file)
        ws = wb.Worksheets(sheet)
        ws.Columns.AutoFit()
        wb.Save()
        excel.Application.Quit()              

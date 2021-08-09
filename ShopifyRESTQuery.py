#!/usr/bin/env python
# -*- coding: utf-8 -*-


"""Main Program
   Name: ShopifyRESTQuery.py

   Calls the relevant functions and prepare de data to
   send to excel files for control
   
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
import time
import threading
from pathlib import Path, PureWindowsPath
from defs import *
from tkinter import *

url = "https://d7434ee9f84db5868a19dd665bc473e2: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@regalatelo-cl.myshopify.com/admin/api/2021-07/"

 
if __name__ == "__main__":

    
      #Esta función retorna un dataframe ordenado de las ordenes
      #generadas (any) desde la fecha dada en el parametro fromDate
      fromDate = "2021-08-02T00:00:00-00:00"
      todate= ""
      orders = get_orders(url,fromDate,todate) 

      #Esta función retorna un dataframe con el resumen de ordenes
      #obtenidas en el dataframe de la funcion get_orders()
      # summaryorders = ordersummary(orders) 

      #La función ProductInventory trae la información del inventario,

      # ProductInventory = get_inventory(url)[0]
      #ProductInventory.rename(columns={'variants.inventory_item_id':'sku'}, inplace=True)
      #La función inventItemIdList devuelve el "id" de inventario unico
      #de cada item
      # inventItemIdList = get_inventory(url)[1]

   
      #La función inventoryItems retorna un dataframe que contiene
      # se puede obtener el valor
      #de costo de cada item de inventario.
      # ItemsCost = get_inventoryItem(inventItemIdList,url)

      #filepath = "C:/Users/Pedro/Documents/Regalatelo/ShopifyScripts/MainProgram/"

      # inventario = ItemsCost.merge(ProductInventory,left_on="id",right_on="variants.inventory_item_id",how="inner")

      #definición de los files paths para open,write and save
      filepath1 = r'C:\Users\Pedro\Documents\Regalatelo\ShopifyScripts\MainProgram\ProductInventory.xlsx'
      filepath2 = r'C:\Users\Pedro\Documents\Regalatelo\ShopifyScripts\MainProgram\InventoryItemsCosts.xlsx'
      filepath3 = r'C:\Users\Pedro\Documents\Regalatelo\ShopifyScripts\MainProgram\orders.xlsx'
      filepath4 = r'C:\Users\Pedro\Documents\Regalatelo\ShopifyScripts\MainProgram\summaryorders.xlsx'


      #LLamada a la función para convertir los dataframes en planillas excel
      # dfToexcel(filepath1,"ProductInventory",inventario) # Aquí se genera la planilla ProductInventory.xlsx en merge con ItemsCost 
      # dfToexcel(filepath2,"ItemsCosts",ItemsCost) # Aquí se genera la planilla InventoryItemsCosts.xlsx
      dfToexcel(filepath3,"orders",orders)  # Aquí se genera la planilla orders.xlsx
      # dfToexcel(filepath4,"summary",summaryorders)  # Aquí se genera la planilla summaryorders.xlsx

      #llamada a la función para setear ancho automático de la columnas en excel
      # setcolumnwidth(filepath1,"ProductInventory")
      # setcolumnwidth(filepath2,"ItemsCosts")
      setcolumnwidth(filepath3,"orders")
      # setcolumnwidth(filepath4,"summary")


      # SkuItemsList = ProductInventory['sku'].values.tolist()

      # print(SkuItemsList)
    
      # valor = (ItemsCost.loc[ItemsCost["sku"] == "SUCSAT01"]).iloc[0]["cost"]

      # print(valor)

      #costo = (valor.iloc[0]["cost"])

      #print(costo)
      
      # print(ProductInventory.head(2))
      # print(ItemsCost.head(2))

      # inventario = ProductInventory.merge(ItemsCost,left_on="sku",right_on = "sku")

      # print(inventario.head())

      #Invent = ProductInventory.sort_values(by=["sku"])
      #Costos = ItemsCost.sort_values(by=["sku"])

      #print(Invent.tail(5))
      #print(Costos.tail(5))


      # 
# -*- coding: utf-8 -*-
"""
Created on Mon Mar  9 16:37:27 2020

@author: usuari
"""

import sys
from gurobipy import *
import ExcelImporter as EI

# Read the excel path from argument
# products, orders = EI.GetData(sys.argv[1])

#Define the days in the week
dies = ["dilluns", "dimarts", "dimecres", "dijous", "divendres"]

orders=[]
contador=0
Num_orders = 10

#while contador < Num_orders:
name="order"+str(contador)
orders.append(name)
contador=contador+1
    
#Define the characteristics of every order
# ordersNumberList = []
# for order in orders:
#     ordersNumberList.append(order.orderNumber)

orders= ["o11","o12","o21","o22","o23","o31"]
# ordersDayList = []
# for order in orders:
#     ordersDayList.append(order.day)
dies_orders=["dilluns","dilluns","dimarts","dimarts","dimarts","dimecres"]

orders_dia=dict(zip(orders,dies_orders))
# orders_dia = dict(zip(ordersNumberList, ordersDayList))

#Define the quantity and the % of lean that we need tot produce in each order
# ordersFatList = []
# for order in orders:
#     ordersFatList.append(order.fatPercentage)
required_lean=[0.5,0.4,0.5,0.4,0.3,0.55]
# ordersMeatKGList = []
# for order in orders:
#     ordersMeatKGList.append(order.meatKg)
total_mass=[1000,1500,2000,2500,500,3000]

#Ceated a diccionary to connect two parameters 
# required_lean_orders=dict(zip(ordersNumberList,ordersFatList))
required_lean_orders=dict(zip(orders,required_lean))   
# total_mass_orders=dict(zip(ordersNumberList,ordersMeatKGList))
total_mass_orders=dict(zip(orders,total_mass))

#Define the suppliers and their products
# suppliers = []
# bodyparts = []
# for product in products:
#     suppliers.append(product.providerName)
#     bodyparts.append(product.productID + '_' + product.name)
# meat_types = [(i,j) for i in suppliers for j in bodyparts]
############################################################
suppliers=["sup1","sup2"]
bodyparts=["A","B","C","D"]
#Define all the procuts that we could buy
meat_types=[(i,j) for i in suppliers for j in bodyparts]
############################################################


#Define the lean % of each meat type
# productFatList = []
# for product in products:
#     productFatList.append(product.fatPercentage)
provided_lean=[0.2,0.3,0.4,0.5,0.25,0.32,0.42,0.55]
meat_types_lean=dict(zip(meat_types,provided_lean))
# meat_types_lean=dict(zip(meat_types,productFatList))

#1 cas original 
# min_quant = []
# for product in products:
#     min_quant.append(product.qtMin)
# min_quant=[250,250,250,250,0,0,0,0]
min_quant=[500,500,500,500,0,0,0,0]
meat_types_minquant=dict(zip(meat_types,min_quant))

# max_quant = []
# for product in products:
#     max_quant.append(product.qtMax)
max_quant=[2500,  2500,  2500,  2000, 1500,  1500,  1500,  1500]
meat_types_maxquant=dict(zip(meat_types,max_quant))

# price = []
# for product in products:
#     price.append(product.price)
price=[10,  12, 14, 16, 12, 15, 16, 20]
meat_types_price=dict(zip(meat_types,price))

storecost=0

#model.update()


#Next, we create a model and the variables. For each product (seven kinds of products) and each time period (month) we will create variables for the amount of which products get manufactured, held and sold. In each month there is an upper limit on the amount of each product that can be sold. This is due to market limitations.
model = Model('Casademont_Oriol')

##################################
############definicio de variables
##################################
amount = model.addVars(suppliers, bodyparts, dies, name="Amount") # quantity bought
use = model.addVars(suppliers, bodyparts, orders, name="Use") # quantity used in orders
buy = model.addVars(suppliers, bodyparts, dies, vtype=GRB.BINARY, name="Buy") # si compro o no compro un tipus de carn
possible_use=model.addVars(suppliers, bodyparts, orders, vtype=GRB.BINARY, name="Pos_use") # si el puc usar per fer una ordre
stock = model.addVars(suppliers, bodyparts, dies, name="Stock") # estoc de cada producte

#sell = model.addVars(time_periods, ub=upper, name="Sell")

model.update()

########################################################################
## CONSTRAINTS #########################################################
########################################################################

# Initialize stock for the first day based on the given parameter
#s.t. Stock_Initialize{(s,b) in MeatTypes}:
#  stock[s,b,0]=initial_stock[s,b];
  
model.addConstrs((amount[sup, body, dia] - (quicksum(use[sup,body,order] 
    for order in orders if orders_dia[order]==dia)) == stock[sup, body,dia] 
    for sup in suppliers for body in bodyparts for dia in dies if dia == dies[0]), 
    name="Initial_Balance")

model.addConstrs((amount[sup, body, dia] + stock[sup,body,dies[dies.index(dia)-1]]
    - (quicksum(use[sup,body,order] for order in orders if orders_dia[order]==dia)) == 
    stock[sup, body,dia] for sup in suppliers for body in bodyparts for dia in dies
    if dia != dies[0]), 
    name="balance")

model.addConstrs((amount[sup, body, dia] <= meat_types_maxquant[sup, body] * buy[sup, body, dia]
    for sup in suppliers for body in bodyparts for dia in dies), 
    name="maximum_Supply")

model.addConstrs((amount[sup, body, dia] >= meat_types_minquant[sup, body] * buy[sup, body, dia]
    for sup in suppliers for body in bodyparts for dia in dies), 
    name="minimum_Supply")

model.addConstrs(((quicksum(use[sup,body,order] for sup in suppliers for body in bodyparts)) == 
    total_mass_orders[order] for order in orders_dia), 
    name="satifer_orders")

model.addConstrs(((quicksum(use[sup,body,order] * meat_types_lean[sup,body] for sup in suppliers for body in bodyparts)) == 
    total_mass_orders[order] * required_lean_orders [order] for order in orders_dia), 
    name="satifer_lean_orders")

#Funcio Objectiu
obj = (quicksum (quicksum(amount[sup,body,dia] for dia in dies) * 
    meat_types_price[sup,body] for sup in suppliers for body in bodyparts) 
    + (quicksum (quicksum(stock[sup,body,dia] for dia in dies) * 
    storecost for sup in suppliers for body in bodyparts)))

model.setObjective(obj, GRB.MINIMIZE)

#Next, we start the optimization and Gurobi tries to find the optimal solution.
model.optimize()

for v in model.getVars():
  if v.X != 0:
      print("%s %f" % (v.Varname, v.X))

#Note: If you want to write your solution to a file, rather than print it to the terminal, you can use the model.write() command. An example implementation is:
model.write('casademont.lp')


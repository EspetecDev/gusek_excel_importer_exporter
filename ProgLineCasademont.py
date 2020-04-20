import sys
from gurobipy import *
import ExcelImporter as EI
import ExcelExporter as EE

excelData = EI.GetData('tfg_oriol.xlsx')
products = excelData[0]
orders = excelData[1]

dies = ["dilluns", "dimarts", "dimecres", "dijous", "divendres"]

orders_number = []
for order in orders:
    orders_number.append(str(order.orderNumber)) #agafem la columna del excel que volem

dies_orders = []
for order in orders:
    dies_orders.append(order.day)

orders_dia = dict(zip(orders_number, dies_orders))


required_lean = []
for order in orders:
    required_lean.append(order.leanPercentage)

total_mass = []
for order in orders:
    total_mass.append(order.meatKg)

required_lean_orders=dict(zip(orders_number,required_lean))

total_mass_orders=dict(zip(orders_number,total_mass))

supplier = []
for product in products:
    supplier.append(product.providerName)

bodypart = []
for product in products:
    bodypart.append(product.productID)
    
suppliers = []
for x in supplier:
     if x not in suppliers: 
            suppliers.append(x)

bodyparts = []
for x in bodypart:
     if x not in bodyparts: 
            bodyparts.append(x)

meat_types=[(i,j) for i in suppliers for j in bodyparts]
    
provided_lean = []
for product in products:
    provided_lean.append(product.leanPercentage)
meat_types_lean=dict(zip(meat_types,provided_lean))

min_quant = []
for product in products:
    min_quant.append(product.qtMin)
meat_types_minquant=dict(zip(meat_types,min_quant))

max_quant = []
for product in products:
    max_quant.append(product.qtMax)
meat_types_maxquant=dict(zip(meat_types,max_quant))

price = []
for product in products:
    price.append(product.price)
meat_types_price=dict(zip(meat_types,price))

storecost=0


model = Model('Casademont_Oriol')

##################################
##VARIABLES#######################
##################################

amount = model.addVars(suppliers, bodyparts, dies, name="Amount") # quantity bought
use = model.addVars(suppliers, bodyparts, orders_number, name="Use") # quantity used in orders
buy = model.addVars(suppliers, bodyparts, dies, vtype=GRB.BINARY, name="Buy") # si compro o no compro un tipus de carn
possible_use=model.addVars(suppliers, bodyparts, orders_number, vtype=GRB.BINARY, name="Pos_use") # si el puc usar per fer una ordre
stock = model.addVars(suppliers, bodyparts, dies, name="Stock") # estoc de cada producte

model.update()


#################################
## CONSTRAINTS ##################
#################################
 
model.addConstrs((amount[sup, body, dia] - (quicksum(use[sup,body,order] 
    for order in orders_number if orders_dia[order]==dia)) == stock[sup, body,dia] 
    for sup in suppliers for body in bodyparts for dia in dies if dia == dies[0]), 
    name="Initial_Balance")

model.addConstrs((amount[sup, body, dia] + stock[sup,body,dies[dies.index(dia)-1]]
    - (quicksum(use[sup,body,order] for order in orders_number if orders_dia[order]==dia)) == 
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

#######################################    
##OBJECTIVE FUNCION####################
#######################################

obj = (quicksum (quicksum(amount[sup,body,dia] for dia in dies) * 
    meat_types_price[sup,body] for sup in suppliers for body in bodyparts) 
    + (quicksum (quicksum(stock[sup,body,dia] for dia in dies) * 
    storecost for sup in suppliers for body in bodyparts)))
model.setObjective(obj, GRB.MINIMIZE)

model.optimize()

#################################
##GENERATED REPORT###############
#################################
#dades = model.getVars()

#for supplier in suppliers:
#    for product_type in bodyparts:
#        for dia in dies:
            



for v in model.getVars():
  if v.X != 0:
      print("%s %f" % (v.Varname, v.X))

model.write('casademont.lp')

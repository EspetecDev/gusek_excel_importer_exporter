import sys, os
import openpyxl

# Una class ens permet definir un "Objecte", en aquest cas un proveidor
# els objectes ens deixen guardar les dade per cada proviedor
class Provider:
	def __init__(self):
		self.name = ''
		self.qtMin = 0.0
		self.qtMax = 0.0
		self.price = 0.0
		self.fatPercentage = 0.0

	def toString(self):
		print('**********************************')
		print('Provider product data: ')
		print('Name: ' + self.name)
		print('Minimum quantity: ' + str(self.qtMin))
		print('Maximum quantity: ' + str(self.qtMax))
		print('Price: ' + str(self.price))
		print('Fat percentage: ' + str(self.fatPercentage))
		print('**********************************')

# i el mateix per cada commanda
class Order:
	def __init__(self):
		self.orderNumber = ''
		self.meatKg = 0
		self.fatPercentage = 0.0
		self.day = 0

	def toString(self):
		print('**********************************')
		print('Order data: ')
		print('Order number: ' + self.orderNumber)
		print('Meat Kg: ' + str(self.meatKg))
		print('Fat percentage: ' + str(self.fatPercentage))
		print('Order day: ' + str(self.day))
		print('**********************************')



# Utility functions that make our life easier

# Get the row and column index for a given cell content
def GetRowColFromTableName(sheet, content):
	for rowidx in list(range(1, sheet.max_row + 1)):
		for colidx in list(range(1, sheet.max_column + 1)):
			if sheet.cell(row = rowidx, column=colidx).value == content:
				# We add 1 to row, since we are looking for the title and we want to start
				# reading the content
				return rowidx+1, colidx 



# Aqui posem els indexos de cada categoria, per saber on hem de llegir de l'excel
# Commanda
##########################
orderNumberCol     	  = 2
meatKgCol 	       	  = 3
fatPercentageCol      = 4
OrderDayCol			  = 5

# Proveidor
###########################
ProveidorCol   	   	  = 7
QuantitatMinimaCol 	  = 8
QuantitatMinimaCol 	  = 9
priceCol			   	  = 10
fatPercentageProveidorCol = 11

# Aqui simplement definim un Array, que es una llista de coses, en aquest cas, de proveidors i una altra de comandes
proveidors = []
commandes = []


# 1. Llegim l'excel i definim les dades

## CONFIGURATION DATA ##
STARTING_PROVIDER_COLUMN_NAME = 'Proveïdor n'
STARTING_ORDER_COLUMN_NAME = 'Numero de comanda'

# Agafem el name de l'excel que volem llegir
filename = 'input_data.xlsx'#sys.argv[1]
print('Parsing file ' + filename + '...')
# El llegim i definim la fulla
file = openpyxl.load_workbook(filename)
sheet = file.active

# 2. Ara agafem les dades
# Lets get where the providers and orders info starts
providersRow, providersCol = GetRowColFromTableName(sheet, STARTING_PROVIDER_COLUMN_NAME)
ordersRow, ordersCol = GetRowColFromTableName(sheet, STARTING_ORDER_COLUMN_NAME)

startParsing = False
for rowidx in list(range(providersRow, sheet.max_row + 1)):
	# Creem un objecte proveidor on guardarem les dades
	tempProvider = Provider()
	tempProvider.name = sheet.cell(row = rowidx, column=providersCol).value
	tempProvider.qtMin = sheet.cell(row = rowidx, column=providersCol+1).value
	tempProvider.qtMax = sheet.cell(row = rowidx, column=providersCol+2).value
	tempProvider.price = sheet.cell(row = rowidx, column=providersCol+3).value
	tempProvider.fatPercentage = sheet.cell(row = rowidx, column=providersCol+4).value
	# Check if the new provider was created successfully
	if tempProvider.name != None:
		proveidors.append(tempProvider)

startParsing = False
# I ara igual amb les commandes
for rowidx in list(range(sheet.max_row + 1)):
	# Comencem a partir de la 4a fila que es on comencen les dades
	if rowidx > 3: 
		tempCommanda = Order()
		tempCommanda.orderNumber = sheet.cell(row = rowidx, column=orderNumberCol).value
		tempCommanda.meatKg = sheet.cell(row = rowidx, column=meatKgCol).value
		tempCommanda.fatPercentage = sheet.cell(row = rowidx, column=fatPercentageCol).value
		tempCommanda.day = sheet.cell(row = rowidx, column=OrderDayCol).value
		if tempCommanda.orderNumber != None:
			print('Nova commanda: ' + tempCommanda.orderNumber)
			commandes.append(tempCommanda)


# Imprimim per pantalla els proveidors que hem llegit
for p in proveidors:
	p.toString()

# I el mateix per les commandes
for c in commandes:
	c.toString()

# 3. Write data to xlsx






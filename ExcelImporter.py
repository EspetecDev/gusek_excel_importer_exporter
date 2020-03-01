import sys, os
import openpyxl

# Una class ens permet definir un "Objecte", en aquest cas un proveidor
# els objectes ens deixen guardar les dade per cada proviedor
class Proveidor:
	def __init__(self):
		self.nom = ''
		self.qtMin = 0.0
		self.qtMax = 0.0
		self.preu = 0.0
		self.percMagre = 0.0

	def toString(self):
		print('**********************************')
		print('Dades del proveidor: ')
		print('Nom: ' + self.nom)
		print('Quantitat Minima: ' + str(self.qtMin))
		print('Quantitat Maxima: ' + str(self.qtMax))
		print('Preu: ' + str(self.preu))
		print('% de Magre: ' + str(self.percMagre))
		print('**********************************')

# i el mateix per cada commanda
class Commanda:
	def __init__(self):
		self.numCommanda = ''
		self.kgCarn = 0
		self.percMagre = 0.0
		self.day = 0

	def toString(self):
		print('**********************************')
		print('Dades de la commanda: ')
		print('Numero de comanda: ' + self.numCommanda)
		print('Kg de Carn: ' + str(self.kgCarn))
		print('% de Magre: ' + str(self.percMagre))
		print('dia de la commanda: ' + str(self.day))
		print('**********************************')


# Aqui posem els indexos de cada categoria, per saber on hem de llegir de l'excel
# Commanda
##########################
NumCommandaCol     	  = 2
KgCarnCol 	       	  = 3
PercMagreCol       	  = 4
OrderDayCol			  = 5

# Proveidor
###########################
ProveidorCol   	   	  = 7
QuantitatMinimaCol 	  = 8
QuantitatMinimaCol 	  = 9
PreuCol			   	  = 10
PercMagreProveidorCol = 11

# Aqui simplement definim un Array, que es una llista de coses, en aquest cas, de proveidors i una altra de comandes
proveidors = []
commandes = []


# 1. Llegim l'excel i definim les dades

# Agafem el nom de l'excel que volem llegir
filename = 'tfg_oriol.xlsx'#sys.argv[1]
print('Parsing file ' + filename)
# El llegim i definim la fulla
file = openpyxl.load_workbook(filename)
sheet = file.active

# 2. Ara agafem les dades
for rowidx in list(range(sheet.max_row + 1)):
	# Comencem a partir de la 4a fila que es on comencen les dades
	if rowidx > 3: 
		# Creem un objecte proveidor on guardarem les dades
		tempProvider = Proveidor()
		tempProvider.nom = sheet.cell(row = rowidx, column=ProveidorCol).value
		tempProvider.qtMin = sheet.cell(row = rowidx, column=QuantitatMinimaCol).value
		tempProvider.qtMax = sheet.cell(row = rowidx, column=QuantitatMinimaCol).value
		tempProvider.preu = sheet.cell(row = rowidx, column=PreuCol).value
		tempProvider.percMagre = sheet.cell(row = rowidx, column=PercMagreProveidorCol).value
		if tempProvider.nom != None:
			print('Nou proveidor: ' + tempProvider.nom)
			proveidors.append(tempProvider)

# I ara igual amb les commandes
for rowidx in list(range(sheet.max_row + 1)):
	# Comencem a partir de la 4a fila que es on comencen les dades
	if rowidx > 3: 
		tempCommanda = Commanda()
		tempCommanda.numCommanda = sheet.cell(row = rowidx, column=NumCommandaCol).value
		tempCommanda.kgCarn = sheet.cell(row = rowidx, column=KgCarnCol).value
		tempCommanda.percMagre = sheet.cell(row = rowidx, column=PercMagreCol).value
		tempCommanda.day = sheet.cell(row = rowidx, column=OrderDayCol).value
		if tempCommanda.numCommanda != None:
			print('Nova commanda: ' + tempCommanda.numCommanda)
			commandes.append(tempCommanda)


# Imprimim per pantalla els proveidors que hem llegit
for p in proveidors:
	p.toString()

# I el mateix per les commandes
for c in commandes:
	c.toString()

# 3. Write data to xlsx
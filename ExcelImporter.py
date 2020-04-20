import openpyxl

# Una class ens permet definir un "Objecte", en aquest cas un proveidor
# els objectes ens deixen guardar les dade per cada proviedor
class Product:
    def __init__(self):
        self.providerName = ''
        self.productID = ''
        self.name = ''
        self.qtMax = 0.0
        self.qtMin = 0.0
        self.leanPercentage = 0.0
        self.price = 0.0

    def toString(self):
        print('**********************************')
        print('Product product data: ')
        print('Provider name: ' + self.providerName)
        print('Product ID: ' + self.productID)
        print('Product Name: ' + self.name)
        print('Maximum quantity: ' + str(self.qtMax))
        print('Minimum quantity: ' + str(self.qtMin))
        print('Lean percentage: ' + str(self.leanPercentage))
        print('Price: ' + str(self.price))
        print('**********************************')

# i el mateix per cada commanda
class Order:
    def __init__(self):
        self.orderNumber = ''
        self.meatKg = 0
        self.leanPercentage = 0.0
        self.day = ' '

    def toString(self):
        print('**********************************')
        print('Order data: ')
        print('Order number: ' + self.orderNumber)
        print('Meat Kg: ' + str(self.meatKg))
        print('Lean percentage: ' + str(self.leanPercentage))
        print('Order day: ' + self.day)
        print('**********************************')





# Get the row and column index for a given cell content
def GetRowColFromTableName(sheet, content):
    indexes = []
    for rowidx in list(range(1, sheet.max_row + 1)):
        for colidx in list(range(1, sheet.max_column + 1)):
            if sheet.cell(row = rowidx, column=colidx).value == content:
                # We add 1 to row, since we are looking for the title and we want to start
                # reading the content
                indexes.append( rowidx+1)
                indexes.append( colidx )
                return indexes

def GetData(filepath):
    
    ## CONFIGURATION DATA ##
    STARTING_PROVIDER_COLUMN_NAME = 'Proveïdor'
    STARTING_ORDER_COLUMN_NAME = 'Numero de comanda'

    # Aqui simplement definim un Array, que es una llista de coses, en aquest cas, de proveidors i una altra de comandes
    products = []
    orders = []

    # Agafem el name de l'excel que volem llegir
    print('Parsing file ' + filepath + '...')
    # El llegim i definim la fulla
    file = openpyxl.load_workbook(filepath)
    sheet = file.get_sheet_by_name("Dad_Ent")

    # Lets get where the products and orders info starts
    productsIndex = GetRowColFromTableName(sheet, STARTING_PROVIDER_COLUMN_NAME)
    productsRow  = productsIndex[0] 
    productsCol = productsIndex[1]

    ordersIndex = GetRowColFromTableName(sheet, STARTING_ORDER_COLUMN_NAME)
    ordersRow = ordersIndex[0]
    ordersCol = ordersIndex[1]

    # Parse products data
    for rowidx in list(range(productsRow, sheet.max_row + 1)):
        # Creem un objecte proveidor on guardarem les dades
        tempProduct = Product()
        tempProduct.providerName = sheet.cell(row = rowidx, column=productsCol).value
        tempProduct.productID = sheet.cell(row = rowidx, column=productsCol+1).value
        tempProduct.name = sheet.cell(row = rowidx, column=productsCol+2).value
        tempProduct.qtMax = sheet.cell(row = rowidx, column=productsCol+3).value
        tempProduct.qtMin = sheet.cell(row = rowidx, column=productsCol+4).value
        tempProduct.leanPercentage = sheet.cell(row = rowidx, column=productsCol+5).value
        tempProduct.price = sheet.cell(row = rowidx, column=productsCol+6).value
        # Check if the new product was created successfully
        if tempProduct.name != None:
            products.append(tempProduct)

    # Parse orders data
    for rowidx in list(range(ordersRow, sheet.max_row + 1)):
        # Comencem a partir de la 4a fila que es on comencen les dades
        tempCommanda = Order()
        tempCommanda.orderNumber = sheet.cell(row = rowidx, column=ordersCol).value
        tempCommanda.meatKg = sheet.cell(row = rowidx, column=ordersCol+1).value
        tempCommanda.leanPercentage = sheet.cell(row = rowidx, column=ordersCol+2).value
        tempCommanda.day = sheet.cell(row = rowidx, column=ordersCol+3).value
        if tempCommanda.orderNumber != None:
            orders.append(tempCommanda)

    return products, orders
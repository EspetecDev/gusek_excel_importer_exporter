import openpyxl

def ExportDataToExcel(self, filename, products, suppliers, days, amount, use, ):
    
    print('Exporting data to ' + filename)
    file = openpyxl.load_workbook(filename)
    sheet = file.get_sheet_by_name("Dad_Sort")
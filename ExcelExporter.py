import openpyxl

def ExportDataToExcel(filename, products, suppliers, days, modelData ):
    
    print('Exporting data to ' + filename)
    # 1. Load excel file and the sheet we will work on
    file = openpyxl.load_workbook(filename)
    sheet = file.get_sheet_by_name("Dad_Sort")

    
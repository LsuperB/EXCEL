from openpyxl import workbook,load_workbook


PN = XX
UN = XX

wb = load_workbook('XX.xlsx')
ws = wb.get_sheet_by_name("XX")
ws = wb.active

wb.save('XX.xlsx')

import openpyxl

wb = openpyxl.load_workbook('/home/ali/Desktop/Talktalk/colored_cell.xlsx')
sheet = wb.get_sheet_by_name("Resources by squad")


color = sheet['J221'].fill.start_color.index

print(color)
import openpyxl
from openpyxl import load_workbook
from openpyxl_image_loader import SheetImageLoader
wb = load_workbook('excelsheet1.xlsx')
st1 = wb['Sheet1']
image_loader = SheetImageLoader(st1)
image = image_loader.get("L3")
image.show()

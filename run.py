import os
from openpyxl import load_workbook, worksheet

def normalize_view(file):
	wb = load_workbook(filename = file)
	for sheet in wb:
		# Put to active cell to A1
		sheet.sheet_view.selection[0].activeCell = 'A1'
		sheet.sheet_view.selection[0].sqref = 'A1'
		sheet.sheet_view.topLeftCell = None

		sheet.sheet_view.tabSelected = False

		# Put the zoom to 100%
		sheet.sheet_view.zoomScale = None
		sheet.sheet_view.zoomScaleNormal = None

	wb.active = 0
	wb.save(file)
	wb.close()
	print(f"{file} OK")

def main():
	path = "."
	excel_files = []
	for root, d, files in os.walk(path):
		for file in files:
			if os.path.splitext(file)[1] in ['.xlsx', '.xls']:
				excel_files.append(os.path.join(root, file))

	print(f"Fichiers à traiter : {len(excel_files)}")
	for file in excel_files:
		normalize_view(file)

main()



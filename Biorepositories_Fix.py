# Clean up spreadsheet to fit C4P format
from openpyxl import load_workbook, cell
from openpyxl.cell import coordinate_from_string, column_index_from_string
import requests
from datetime import datetime, date, time

wb = load_workbook('biorepositories_from_website.xlsx')
ws = wb.get_active_sheet()

START_ROW = 2
END_ROW = ws.get_highest_row()
GR_URL_COL = 'A'
GR_INSTIT_NAME_COL = 'B'
GR_REPO_NAME_COL = 'C'

class Entry(self, name):
	self.name = name
	#institutionName
	#repoName
	#URL
	#desc
	#contact

for row in ws.range('%s%s:%s%s'%(GR_INSTIT_NAME_COL, START_ROW, GR_INSTIT_NAME_COL, END_ROW)):
	for cell in row:
		if

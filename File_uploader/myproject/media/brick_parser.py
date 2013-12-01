from xlrd import open_workbook

wb = open_workbook('BOM_export.xlsx')
sheet = wb.sheet_by_index(0)

idsRaw =  sheet.col(3)
colorsRaw = sheet.col(4)
colorswHF = []
idswHF = []

for rows in idsRaw :
	if isinstance(rows.value,float):
		idswHF.append(str(int(rows.value)))
	else:
		idswHF.append(rows.value)
for rows in colorsRaw : 
	stripped =  rows.value.split(" ")
	colorswHF.append(stripped[0].encode('ascii','ignore'))

#Lets go ahead and strip the header and footer off
#One of a million ways to do this
ids = idswHF[1:(len(idswHF)-1)]
colors = colorswHF[1:(len(colorswHF)-1)]
print ids
print colors
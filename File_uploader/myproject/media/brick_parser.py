from xlrd import open_workbook
from lxml import etree as ET

def parse_bricks(excel_sheet):
	wb = open_workbook(excel_sheet)
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
	print ids[0]
	print colors
	##Build out the XML
	root = ET.Element("INVENTORY")
	for i in range(len(ids)):
		item = ET.SubElement(root,"ITEM")
		itemType = ET.SubElement(item,"ITEMTYPE")
		itemType.text = "P"
		itemID = ET.SubElement(item,"ITEMID")
		itemID.text = ids[i]
		ItemColor = ET.SubElement(item,"COLOR")
		ItemColor.text = colors[i]
	tree = ET.ElementTree(root)
	tree.write("output.xml", pretty_print=True)

#def build_xml():
#	root = ET.Element("INVENTORY")
#	for i in ids:
#		item = ET.SubElement(root,"ITEM")
#		itemType = ET.SubElement(item,"P")
#		itemID = ET.SubElement(item,ids[i])
#		ItemColor = ET.Subelement(item,colors[i])

parse_bricks('BOM_export.xlsx')
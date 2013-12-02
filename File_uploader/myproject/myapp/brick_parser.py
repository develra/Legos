from xlrd import open_workbook
from lxml import etree as ET

#Takes a BOM export from LDD, exports an .xml sheet for bricklink
def l2b_colors(ldd_color):
	f = open('Colors.txt',"r")
	for line in f:
		if (ldd_color==line.split("	")[0]):
			return line.split("	")[2]
	f.close()

def l2b_ids(ldd_id, ldd_color):
	tree = ET.parse("BLPartAndColorCodes.xml")
	root = tree.getroot()
	for item in root[1]:
		if (item[3].text==ldd_id):
			return item[1].text
	return ldd_id.rstrip(ldd_color)


def parse_bricks(excel_sheet):
	wb = open_workbook(excel_sheet)
	sheet = wb.sheet_by_index(0)

	idsRaw =  sheet.col(0)
	colorsRaw = sheet.col(4)
	quantRaw = sheet.col(5)
	colorswHF = []
	idswHF = []
	quantwHF = []

	for rows in idsRaw:
		if isinstance(rows.value,float):
			idswHF.append(str(int(rows.value)))
		else:
			idswHF.append(rows.value)
	for rows in colorsRaw: 
		stripped =  rows.value.split(" ")
		colorswHF.append(stripped[0].encode('ascii','ignore'))
	for rows in quantRaw:
		if isinstance(rows.value,float):
			quantwHF.append(str(int(rows.value)))
		else:
			quantwHF.append(rows.value)
	#Lets go ahead and strip the header and footer off
	#One of a million ways to do this
	ids = idswHF[1:(len(idswHF)-1)]
	colors = colorswHF[1:(len(colorswHF)-1)]
	quant = quantwHF[1:(len(quantwHF)-1)]
	print ids[0]
	print colors
	print quant
	##Build out the XML
	root = ET.Element("INVENTORY")
	for i in range(len(ids)):
		item = ET.SubElement(root,"ITEM")
		itemType = ET.SubElement(item,"ITEMTYPE")
		itemType.text = "P"
		itemID = ET.SubElement(item,"ITEMID")
		itemID.text = l2b_ids(ids[i],colors[i])
		ItemColor = ET.SubElement(item,"COLOR")
		#need to convert LDD colors 2 bricklink colors
		ItemColor.text = l2b_colors(colors[i])
		itemQuant = ET.SubElement(item,"MINQTY")
		itemQuant.text = quant[i]
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
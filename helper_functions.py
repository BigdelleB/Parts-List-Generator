import math
from openpyxl import Workbook
from openpyxl.styles import colors
from openpyxl.styles import Font, Color
from openpyxl import load_workbook

'''
Input the file path for the bom_list 
Return the categories to put into the excel file

SHOULD BE [ITEM, QTY, REFERENCE, DESCRIPTION, MANUFACTURER, MANUFACTURER P/N, MIC]
'''
def get_columns(BOM_list):
	columns = ["ITEM", "QTY", "REFERENCE", "DESCRIPTION", "MANUFACTURER","MANUFACTURER P/N", "MIC"]
	return columns;
  
'''
Return an array of every part in the parts list.
EVERY ROW OF PARTS HAS ["ITEM", "QTY", "REFERENCE", "DESCRIPTION", "MANUFACTURER","MANUFACTURER P/N"]
	w/o the MIC part, ignore that for now, let the User fill it in because I dont see any category for that on the parts list
'''
def get_parts_list(BOM_list):
	parts = []
	ITEM=""
	QTY=""
	REFERENCE=""
	DESCRIPTION=""
	MANUFACTURER=""
	MANUFACTURER_PN=""
	found = False

	#go through the BOM list, line by line
	with open(BOM_list) as fp:
			line  = "temp"
			while not found and line:
			#get line
				line = fp.readline()
				#find when part numbers start being referenced
				if "Item" in line.split():
					found = True
					fp.readline()
					fp.readline()
					line = fp.readline()

			while found and line:
				#this is the parts line
				line = line.split()

				#get the first 3 fields, "item, qty and description"
				ITEM = line.pop(0)
				QTY = line.pop(0)
				REFERENCE = line.pop(0)
				#get manufacturer P/N, these are easy because theyre one string
				MANUFACTURER_PN = line.pop(len(line)-1)
				MANUFACTURER=line.pop(len(line)-1)

				seperator = " "
				DESCRIPTION = seperator.join(line)

				parts.append([ITEM,QTY,REFERENCE,DESCRIPTION,MANUFACTURER_PN,MANUFACTURER])
				line = fp.readline()

			
	return parts



'''
Return the number of items in the list
'''
def get_num_items(BOM_list):
	with open(BOM_list) as f:
		for line in f:
			pass
		last_line = line

	last_line = last_line.split()

	return last_line[0]



def main(BOM_list, file_name):
	#define an excel file, this is the output
	wb = load_workbook(filename = 'Cover Page.xlsx')
	wb.font = colors.BLACK



	categories = get_columns(BOM_list)
	parts = get_parts_list(BOM_list)
	numparts = float(get_num_items(BOM_list))

	i=-1;


	#make sure each parts page has max of 27 parts
	num_pages = math.ceil(numparts/27);

	for page in range(num_pages):
		#create a new page
		wb.create_sheet("Page "+ str(1+page))
		ws = wb["Page "+ str(1+page)]
		ws.append(categories)

		if(page < num_pages-1):
			for i in range(page*26,(page+1)*26):
				#print(str(page)+" " +str(i))

				row = parts[i]
				ws.append(row)

		else:
			for i in range(page*26,len(parts)):
				row = parts[i]
				ws.append(row)




	wb.save("PL"+file_name+"-0001_.xlsx")














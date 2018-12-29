import fnmatch
import os
import xlrd 
import datetime
import csv
import numpy as np

DESIGNATOR_START_ROW = 11
LIBREF_COLUMN = 0
DESIGNATOR_COLUMN = 1
DIGIKEY_PN_COLUMN = 5
QUANTITY_COLUMN = 7


common_components_part_numbers = np.genfromtxt('/Users/Taiping/Desktop/common_components.csv', delimiter=',', dtype=str, usecols=1, autostrip=True)

def include_component(libref, quantity):
	prompt = "BOM Contains {0}x {1}. \nPress Enter to order it. \nPress any other key to skip it.".format(int(quantity), str(libref)) 
	user_input = raw_input(prompt)
	if user_input == "":
		print "Ordered...\n"
		return True
	else:
		print "Skipping...\n"
		return False

def remove_non_ascii(text): 
	return "".join(i for i in text if ord(i)<128)


matches = []
array_file_name = []
array_root_directories = []

current_directory = os.getcwd()

for root, dirnames, filenames in os.walk(current_directory):
    for excel_file in fnmatch.filter(filenames, '*.xlsx'):
    	if "~$" not in excel_file:
        	array_file_name.append(excel_file)
        	array_root_directories.append(root)

selection_number = 0

for file in array_file_name:
	print str(selection_number) + ': ' + str(file)
	selection_number = selection_number + 1
print "\n"


selection_number = int(raw_input('Choose which file to parse...   '))

filename = "{0}/{1}".format(array_root_directories[selection_number], array_file_name[selection_number])

workbook = xlrd.open_workbook(filename)
worksheet = workbook.sheet_by_name("BOM Report") # We need to read the data 
num_rows = worksheet.nrows #Number of Rows
num_cols = worksheet.ncols #Number of Columns

designator_prefix = str(raw_input('What should the prefix be?  '))
print "\n"

array_librefs = []
array_designators = []
array_digikey_part_numbers = []
array_quantity = []

for curr_row in range(DESIGNATOR_START_ROW, num_rows-1):
	lib_ref = worksheet.cell_value(curr_row, LIBREF_COLUMN)

	if "Test Point" in lib_ref:
		continue

	quantity = int(worksheet.cell_value(curr_row, QUANTITY_COLUMN))
	digikey_pn = worksheet.cell_value(curr_row, DIGIKEY_PN_COLUMN)

	# Check for connectors
	if "CONN" in lib_ref:
		if not include_component(lib_ref, quantity):
			continue

	if digikey_pn in common_components_part_numbers:
		if not include_component(lib_ref, quantity):
			continue

	array_librefs.append(lib_ref)

	designator = str(worksheet.cell_value(curr_row, DESIGNATOR_COLUMN))
	designator = "{0}_{1}".format(designator_prefix, designator)

	if len(str(designator)) > 48:
		designator = ''.join(designator.split())

	array_designators.append(designator)

	array_digikey_part_numbers.append(digikey_pn)

	array_quantity.append(quantity)

print array_quantity

desktop_path = os.path.expanduser("~/Desktop") 
order_file = "{0}/Digikey_{1}.csv".format(desktop_path, datetime.datetime.now().strftime('%Y_%m_%d')) 

print "Writing values to {0}".format(order_file)
print "\n"

if os.path.isfile(order_file):
	file = open(order_file, 'ab')
	writer = csv.writer(file)

else:
	file = open(order_file, 'wb')
	writer = csv.writer(file)
	writer.writerow(("LibRefs", "Designators", "Supplier Part Numbers", "Quantity"))


number_of_rows = num_rows - DESIGNATOR_START_ROW - 1

order_quantity = int(raw_input("Enter Quantity to be Ordered as an integer...    "))

for row in range(0, len(array_librefs)):
	try:
		writer.writerow((array_librefs[row], array_designators[row], array_digikey_part_numbers[row], int(array_quantity[row]) * order_quantity))
	except UnicodeEncodeError:
		lib_ref = remove_non_ascii(array_librefs[row])
		part_number = remove_non_ascii(array_digikey_part_numbers[row])
		writer.writerow((lib_ref, array_designators[row], part_number, int(array_quantity[row]) * order_quantity))

print "Finished writing values to " + order_file

file.close()
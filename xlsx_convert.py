import xlrd
import datetime
import re

### Setup for Reading .xlsx ###
# 讀取檔案
read_xlsx_file_name = 'I18N_iOS_string_table_20200810_miny.xlsx'
# 專案內表示文字之欄位
column_expression = 0 
# 欲轉換語系之欄位
column_translate = 1

### Output Header ###
header_1_export_file_name = 'Localizable.strings'
header_2_from_where = 'from xlsx_convert.py' 
header_3_created_by = 'auto generated'
header_4_create_date = datetime.datetime.now().strftime('%Y/%-m/%-d')
header_5_create_year = datetime.datetime.now().strftime('%Y')
header_6_copyright = 'JohnsonTechInc.'
output = '//\n//  %s\n//  %s\n//\n//  created by %s on %s.\n//  Copyright © %s %s All rights reserved.\n//\n\n' % (header_1_export_file_name, header_2_from_where, header_3_created_by, header_4_create_date, header_5_create_year, header_6_copyright)

### Data Process ###
print('=== START PROCESS, xlsx File: %s  ===' % read_xlsx_file_name)
print('\n')
print('Expression Column: %d' % column_expression)
print('Translate Column: %d' % column_translate)
print('\n')

workbook = xlrd.open_workbook(read_xlsx_file_name)
sheet = workbook.sheet_by_index(0)
last_match_str = ''

for row in range(column_expression, sheet.nrows):

	### TODO: Read Localized Target to CI ###
	if row == 0: continue
	
	str_expression = sheet.cell_value(row,column_expression)

	# prevent from wrong value, wrong translate to 'N/A'
	# 5 means error, 6 means empty
	cell_type = sheet.cell_type(row, column_translate)

	# skip row
	if row == 1: 
		print("skip row: %d, cell_type: %d, key: %s" % (row, cell_type, sheet.cell_value(row - 1,0)))

	# check cell type, refer to https://xlrd.readthedocs.io/en/latest/api.html#xlrd.sheet.Cell
	# type = 0 : empty
	# type = 5 : error
	# type = 6 : blank
	if cell_type != 0 and cell_type != 5 and cell_type != 6:
		
		str_translate = sheet.cell_value(row,column_translate)
		
		# only grab type = 1 : unicode string 
		# find " character, add \
		if cell_type == 1:
			str_translate = re.sub(r'(")', r'\\"', str_translate)

	else:
		print("row: %d empty/error/blank, cell_type: %d, key: %s" % (row, cell_type, sheet.cell_value(row - 1,0)))
		continue
		# if wanna cetrain text write in
		# str_translate = 'N/A'

	export_line = '%s = "%s";' % (str_expression, str_translate)

	# detect localized expression split string 1st, 2nd same
	# if true add notes
	if row + 1 < sheet.nrows:

		str_next_expression = sheet.cell_value(row + 1,column_expression)

		# split results
		current_split_results = str_expression.split('.', 3)
		next_split_results = str_next_expression.split('.', 3)

		# output condition: less than three phrases
		if len(current_split_results) < 3: 
			last_match_str = ''
			output = output + '\n' + export_line + '\n'
			continue

		phrases_current = ''
		pharses_next = ''
		count = 0

		# get current first 2 pharses
		for r in current_split_results:
			if count > 1: # read first 2
				break 
			phrases_current = phrases_current + r
			count = count + 1
		
		# if match with last matched string, continue(skip)
		if phrases_current == last_match_str:
			# output condition: already matched
			output = output + export_line + '\n'
			continue
		# if not, clear last matched string
		else:
			last_match_str = ''

		# get next first 2 pharses
		count = 0
		for r in next_split_results:
			if count > 1: # read first 2
				break
			pharses_next = pharses_next + r
			count = count + 1

		# current & next matched, add notes
		if phrases_current == pharses_next:
			export_line = '\n//%s\n' % (phrases_current) + export_line
			last_match_str = phrases_current

	# output condition: normal
	output = output + export_line + '\n'

### Output File ###
text_file = open(header_1_export_file_name, "w")
text_file.write(output)
text_file.close()
print('\n=== COMPLETE, Generated file: %s ===' % header_1_export_file_name)
# print('' )




import os
import re
import subprocess
import glob
import xlsxwriter
import datetime


# *** Tool executing procedure ***
# Enter the JHOVE bat path in the jhove.ini file
# execute the tool and enter the file path
# Validation xlsx will be created in the file path

current_date = datetime.datetime.now().date()

limitation_date = datetime.date(2023, 12, 29)

if current_date > limitation_date:
    print("Date limit exceeded. Program cannot run.")
    exit()

print("\n *** TIFF - JHOVE Validation *** \n")

filepath = input(" Enter the File path: ")

text_file = filepath + "/" + "error.log"

jhove = "jhove.ini"

# check the exiftool file present
if os.path.exists(jhove):
	pass
else:
	print("\n jhove.ini tool is missing")
	f = open(text_file, "a+")
	f.write(str("jhove.ini tool is missing\n"))
	f.close()
	exit()

#reading the tag.ini file
fo = open(jhove, "r+", encoding="utf-8")
val1 = fo.read()
text = re.search(r"<path>(.*?)</path>", str(val1))
val2 = str("\"") + text.group(1) + str("\"")
val3 = text.group(1)

if os.path.exists(val3):
	pass
else:
	print("\n JHOVE path is not correct")
	f = open(text_file, "a+")
	f.write(str("JHOVE path is not correct\n"))
	f.close()
	exit()

excelfile = filepath + "/" + 'Validation_log.xlsx'
workbook = xlsxwriter.Workbook(excelfile)
worksheet = workbook.add_worksheet()

worksheet.hide_gridlines(2)

cell_format_border = workbook.add_format({
    'border': 1,  # 1 indicates a thin border
    'align': 'center',  # Optional: center align text
    'valign': 'vcenter',  # Optional: center align vertically
    'bold': True
})

worksheet.write('A1', 'TIFF File Name', cell_format_border)
worksheet.write('B1', 'Status', cell_format_border)
worksheet.write('C1', 'TIFF Version', cell_format_border)
worksheet.write('D1', 'Compression', cell_format_border)
worksheet.write('E1', 'Width', cell_format_border)
worksheet.write('F1', 'Height', cell_format_border)
worksheet.write('G1', 'Color Mode', cell_format_border)
worksheet.write('H1', 'ICC Profile', cell_format_border)
worksheet.write('I1', 'Date Time', cell_format_border)
worksheet.write('J1', 'Artist', cell_format_border)
worksheet.write('K1', 'Scanner Manufacture', cell_format_border)
worksheet.write('L1', 'Scanner Model', cell_format_border)
worksheet.write('M1', 'Scanner Software', cell_format_border)
worksheet.write('N1', 'Orientation', cell_format_border)
worksheet.write('O1', 'Resolution (x, y)', cell_format_border)
worksheet.write('P1', 'Bits Per Sample', cell_format_border)
worksheet.write('Q1', 'Sample Per Pixel', cell_format_border)

row = 1

for fname in glob.glob(filepath + "/" + "*.tif"):
	input_file = str("\"") + fname + str("\"")
	output1 = filepath + "/" + "out.xml"
	name = os.path.basename(fname)
	print(name)
	splitname = os.path.splitext(name)[0]
	jp2_filename = str("\"") + filepath + "/" + splitname + ".jp2" + str("\"")
	conversion = val2 + " " + "-m tiff-hul"+ " " + input_file + " "  + "-h XML" + " " + "-o " + output1
	subprocess.call(conversion)
	xdpi = None
	ydpi = None
	resol = None
	fo = open(output1, "r", encoding="utf-8")
	content = fo.read()
	content = re.sub(r'\n\s*', '', content)  # Remove spaces after newline characters
	content = re.sub(r'\n', '', content)  # Remove newline characters
	content = re.sub(r'</mix:bitsPerSampleValue><mix:bitsPerSampleValue>', ', ', content)
	content = re.sub(r'<mix:bitsPerSampleValue>', '', content)
	content = re.sub(r'</mix:bitsPerSampleValue>', '', content)
	content = re.sub(r'</mix:bitsPerSampleValue>', '', content)
	content = re.sub(r'</mix:numerator></mix:xSamplingFrequency>', '</mix:numerator><mix:denominator>1</mix:denominator></mix:xSamplingFrequency>', content)
	content = re.sub(r'</mix:numerator></mix:ySamplingFrequency>', '</mix:numerator><mix:denominator>1</mix:denominator></mix:ySamplingFrequency>', content)
	content = re.sub(r'<mix:xSamplingFrequency><mix:numerator>', '<xdpi>', content)
	content = re.sub(r'</mix:denominator></mix:xSamplingFrequency>', '</xdpi>', content)
	content = re.sub(r'<mix:ySamplingFrequency><mix:numerator>', '<ydpi>', content)
	content = re.sub(r'</mix:denominator></mix:ySamplingFrequency>', '</ydpi>', content)
	content = re.sub(r'</mix:numerator><mix:denominator>', '<divide/>', content)

	tag1 = re.search(r"<status>(.*?)</status>", str(content))
	status = tag1.group(1) if tag1 else "Status not found - File not correct"
	tag2 = re.search(r"<version>(.*?)</version>", str(content))
	ver = tag2.group(1) if tag2 else "Version not found - File not correct"
	tag3 = re.search(r"<mix:compressionScheme>(.*?)</mix:compressionScheme>", str(content))
	comp = tag3.group(1) if tag3 else "Compression not found - File not correct"
	tag4 = re.search(r"<mix:imageWidth>(.*?)</mix:imageWidth>", str(content))
	wid = tag4.group(1) if tag4 else ""
	tag5 = re.search(r"<mix:imageHeight>(.*?)</mix:imageHeight>", str(content))
	hei = tag5.group(1) if tag5 else ""
	tag6 = re.search(r"<mix:colorSpace>(.*?)</mix:colorSpace>", str(content))
	tag6_1 = tag6.group(1) if tag5 else ""
	if tag6_1 == "WhiteIsZero":
		color = "Black and White"
	elif tag6_1 == "BlackIsZero":
		color = "Grayscale"
	else:
		color = tag6_1
	tag7 = re.search(r"<mix:iccProfileName>(.*?)</mix:iccProfileName>", str(content))
	profile = tag7.group(1) if tag7 else ""
	tag8 = re.search(r"<mix:dateTimeCreated>(.*?)</mix:dateTimeCreated>", str(content))
	datetime = tag8.group(1) if tag8 else ""
	tag9 = re.search(r"<mix:imageProducer>(.*?)</mix:imageProducer>", str(content))
	artist = tag9.group(1) if tag9 else ""
	tag10 = re.search(r"<mix:scannerManufacturer>(.*?)</mix:scannerManufacturer>", str(content))
	scan = tag10.group(1) if tag10 else ""
	tag11 = re.search(r"<mix:scannerModelName>(.*?)</mix:scannerModelName>", str(content))
	scanmod = tag11.group(1) if tag11 else ""
	tag12 = re.search(r"<mix:scanningSoftwareName>(.*?)</mix:scanningSoftwareName>", str(content))
	scansof = tag12.group(1) if tag12 else ""
	tag13 = re.search(r"<mix:orientation>(.*?)</mix:orientation>", str(content))
	ori = tag13.group(1) if tag13 else ""

	tag14_1 = re.search(r"<xdpi>(.*?)<divide/>(.*?)</xdpi>", str(content))
	tag14_2 = tag14_1.group(1) if tag14_1 else ""
	tag14_3 = tag14_1.group(2) if tag14_1 else ""
	try:
		tag14_4 = float(tag14_2)
		tag14_5 = float(tag14_3)
		if tag14_5 == 0:
			print("Error: Division by zero")
		else:
			tag14_6 = tag14_4 / tag14_5
			xdpi = int(tag14_6)
			print(xdpi)
	except ValueError:
		print("Resolution not found")

	tag14_7 = re.search(r"<ydpi>(.*?)<divide/>(.*?)</ydpi>", str(content))
	tag14_8 = tag14_7.group(1) if tag14_7 else ""
	tag14_9 = tag14_7.group(2) if tag14_7 else ""
	try:
		tag14_10 = float(tag14_8)
		tag14_11 = float(tag14_9)
		if tag14_11 == 0:
			print("Error: Division by zero")
		else:
			tag14_12 = tag14_10 / tag14_11
			ydpi = int(tag14_12)
	except ValueError:
		print("Resolution not found")
	if xdpi is not None and ydpi is not None:
		resol = f"{xdpi}, {ydpi}"
	tag15 = re.search(r"<mix:numerator>(.*?)</mix:numerator>", str(content))
	ydpi = tag15.group(1) if tag15 else ""
	tag16 = re.search(r"<mix:BitsPerSample>(.*?)<mix:bitsPerSampleUnit>", str(content))
	bits = tag16.group(1) if tag16 else ""
	tag17 = re.search(r"<mix:samplesPerPixel>(.*?)</mix:samplesPerPixel>", str(content))
	sampleper = tag17.group(1) if tag17 else ""



	cell_format_status = workbook.add_format({'border': 1})
	cell_format_ver = workbook.add_format({'border': 1})
	cell_format_bor = workbook.add_format({'border': 1})

	worksheet.write(row, 0, name, cell_format_bor)
	if status != "Well-Formed and valid":
		cell_format_status.set_bg_color('yellow')
		worksheet.write(row, 1, status, cell_format_status)
	else:
		worksheet.write(row, 1, status, cell_format_status)
	if ver != "6.0":
		cell_format_ver.set_bg_color('yellow')
		worksheet.write(row, 2, ver, cell_format_ver)
	else:
		worksheet.write(row, 2, ver, cell_format_ver)
	worksheet.write(row, 3, comp, cell_format_bor)
	worksheet.write(row, 4, wid, cell_format_bor)
	worksheet.write(row, 5, hei, cell_format_bor)
	worksheet.write(row, 6, color, cell_format_bor)
	worksheet.write(row, 7, profile, cell_format_bor)
	worksheet.write(row, 8, datetime, cell_format_bor)
	worksheet.write(row, 9, artist, cell_format_bor)
	worksheet.write(row, 10, scan, cell_format_bor)
	worksheet.write(row, 11, scanmod, cell_format_bor)
	worksheet.write(row, 12, scansof, cell_format_bor)
	worksheet.write(row, 13, ori, cell_format_bor)
	worksheet.write(row, 14, resol, cell_format_bor)
	worksheet.write(row, 15, bits, cell_format_bor)
	worksheet.write(row, 16, sampleper, cell_format_bor)

	row += 1

	fo.close()
	os.remove(output1)

worksheet.autofit()
workbook.close()

print("\n*** JHOVE Validation Completed ***")

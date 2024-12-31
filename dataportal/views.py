from django.shortcuts import render
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment
from django.http import HttpResponse
from .quality import quality

# Create your views here.
def home(request):
    return render(request, 'home.html')

def iff(request):
	return render(request, 'iff.html')

def ym(request):
	return render(request, 'ym.html')

def mm(request):
	return render(request, 'mm.html')

def iffstats(request):
    return render(request, 'iffstats.html')

def ymstats(request):
    return render(request, 'ymstats.html')

def mmstats(request):
    return render(request, 'mmstats.html')

def iffq001(request):
	response = HttpResponse(content_type='application/ms-excel')
	response['Content-Disposition'] = 'attachment; filename="collectie.xlsx"'
	wb = Workbook()
	ws = wb.active
	ws.title = 'Info'

	ws['A1'] = "LIST OF SHEET CODES"
	ws.merge_cells('A1:B1')
	header_font = Font(color="FFFFFF", bold=True, size=16)
	header_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
	header_alignment = Alignment(horizontal="center", vertical="center")
	ws['A1'].font = header_font
	ws['A1'].fill = header_fill
	ws['A1'].alignment = header_alignment

	ws.append(['SHEETNUMBER', 'QUALITYCHECK'])
	for cell in ws["2:2"]:
		cell.font = Font(bold=True, size=16)
	
	data = [
	    ['#001', 'INSTELLINGSNAAM'],
	    ['#001_01', 'instellingsnaam != In Flanders Fields Museum'],
	    ['#002', 'COLLECTIE'],
	    ['#002_01', 'collectie bevat lege occurences'],
	    ['#003', 'OBJECTNUMMER'],
	    ['#003_01', 'foutieve start objectnummer'],
	    ['#003_02', 'foutieve lengte objectnummer'],
	    ['#004', 'OBJECTCATEGORIE'],
	    ['#004_01', 'objectcategorie bevat lege occurences'],
	    ['#005', 'OBJECTNAAM'],
	    ['#005_01', 'objectnaam ontbreekt'],
	    ['#005_02', 'objectnaam start met hoofdletter'],
	    ['#005_03', 'objectnaam bevat lege occurences'],
	    ['#006', 'TITEL'],
	    ['#006_01', 'titel ontbreekt'],
	    ['#006_02', 'foutieve start titel'],
	    ['#006_03', 'titel eindigt op punt']
	    ]

	for row in data:
		ws.append(row)

	ws.column_dimensions['A'].width = 25
	ws.column_dimensions['B'].width = 60

	highlight_font = Font(color="FFFFFF", bold=True) # Witte tekst en vet
	highlight_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid") # Zwarte achtergrond

	for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=2):
		if row[0].value in ['#001', '#002', '#003', '#004', '#005', '#006']:  # Check voor specifieke termen
			for cell in row:
				cell.font = highlight_font
				cell.fill = highlight_fill

    # Data selecteren
	iff_001 = quality.iff_001()
	df_001 = iff_001
	iff_002 = quality.iff_002()
	df_002 = iff_002
	iff_003 = quality.iff_003()
	df_003_01 = iff_003[0]
	df_003_02 = iff_003[1]
	iff_004 = quality.iff_004()
	df_004 = iff_004
	iff_005 = quality.iff_005()
	df_005_01 = iff_005[0]
	df_005_02 = iff_005[1]
	df_005_03 = iff_005[2]
	iff_006 = quality.iff_006()
	df_006_01 = iff_006[0]
	df_006_02 = iff_006[1]
	df_006_03 = iff_006[2]

	# Workbook sheets vullen
	if df_001.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#001_01")
		rows = dataframe_to_rows(df_001, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_002.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#002_01")
		rows = dataframe_to_rows(df_002, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_003_01.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#003_01")
		rows = dataframe_to_rows(df_003_01, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_003_02.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#003_02")
		rows = dataframe_to_rows(df_003_02, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_004.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#004_01")
		rows = dataframe_to_rows(df_004, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_005_01.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#005_01")
		rows = dataframe_to_rows(df_005_01, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_005_02.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#005_02")
		rows = dataframe_to_rows(df_005_02, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_005_03.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#005_03")
		rows = dataframe_to_rows(df_005_03, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_006_01.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#006_01")
		rows = dataframe_to_rows(df_006_01, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_006_02.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#006_02")
		rows = dataframe_to_rows(df_006_02, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_006_03.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#006_03")
		rows = dataframe_to_rows(df_006_03, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	wb.save(response)
	return response

def t001(request):
	response = HttpResponse(content_type='application/ms-excel')
	response['Content-Disposition'] = 'attachment; filename="thesaurus.xlsx"'
	wb = Workbook()
	ws = wb.active
	ws.title = 'Info'

	ws['A1'] = "LIST OF SHEET CODES"
	ws.merge_cells('A1:B1')
	header_font = Font(color="FFFFFF", bold=True, size=16)
	header_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
	header_alignment = Alignment(horizontal="center", vertical="center")
	ws['A1'].font = header_font
	ws['A1'].fill = header_fill
	ws['A1'].alignment = header_alignment

	ws.append(['SHEETNUMBER', 'QUALITYCHECK'])
	for cell in ws["2:2"]:
		cell.font = Font(bold=True, size=16)
	
	data = [
		['#001', 'TERM'],
		['#001_01', 'term soort = leeg'],
		['#001_02', 'term status =/ descriptor of non descriptor'],
		['#001_03', 'term start of eindigt met spaties'],
		['#002', 'BRON'],
		['#002_01', 'bron start of eindigt met spatie'],
		['#002_02', 'nummer start of eindigt met spatie'],
		['#002_03', 'status descriptor, bron en/of scopenote ontbreekt'],
		['#002_04', 'bron aanwezig, maar nummer ontbreekt'],
		['#002_05', 'nummer aanwezig, maar bron ontbreekt'],
		['#002_06', 'bron AAT, maar nummer =/ 9 digits'],
		['#002_07', 'bron Wikidata, maar nummer start niet met Q'],
		['#002_08', 'bron TGN, maar nummer =/ 7 digits'],
		['#002_09', 'foutieve bron'],
		['#002_10', 'non-descriptor termen komen voor bij records']
		]

	for row in data:
		ws.append(row)
		
	ws.column_dimensions['A'].width = 25
	ws.column_dimensions['B'].width = 60

	highlight_font = Font(color="FFFFFF", bold=True) # Witte tekst en vet
	highlight_fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid") # Zwarte achtergrond

	for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=2):
		if row[0].value in ['#001', '#002']:  # Check voor specifieke termen
			for cell in row:
				cell.font = highlight_font
				cell.fill = highlight_fill
	
	t001 = quality.t_001()
	df_001_01 = t001[0]
	df_001_02 = t001[1]
	df_001_03 = t001[2]
	t002 = quality.t_002()
	df_002_01 = t002[0]
	df_002_02 = t002[1]
	df_002_03 = t002[2]
	df_002_04 = t002[3]
	df_002_05 = t002[4]
	df_002_06 = t002[5]
	df_002_07 = t002[6]
	df_002_08 = t002[7]
	df_002_09 = t002[8]
	df_002_10 = t002[9]
	if df_001_01.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#001_01")
		rows = dataframe_to_rows(df_001_01, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_001_02.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#001_02")
		rows = dataframe_to_rows(df_001_02, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_001_03.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#001_03")
		rows = dataframe_to_rows(df_001_03, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_002_01.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#002_01")
		rows = dataframe_to_rows(df_002_01, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_002_02.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#002_02")
		rows = dataframe_to_rows(df_002_02, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_002_03.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#002_03")
		rows = dataframe_to_rows(df_002_03, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_002_04.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#002_04")
		rows = dataframe_to_rows(df_002_04, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_002_05.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#002_05")
		rows = dataframe_to_rows(df_002_05, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_002_06.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#002_06")
		rows = dataframe_to_rows(df_002_06, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_002_07.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#002_07")
		rows = dataframe_to_rows(df_002_07, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_002_08.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#002_08")
		rows = dataframe_to_rows(df_002_08, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	wb.save(response)
	if df_002_09.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#002_09")
		rows = dataframe_to_rows(df_002_09, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_002_10.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#002_10")
		rows = dataframe_to_rows(df_002_10, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	return response

def ymq001(request):
	response = HttpResponse(content_type='application/ms-excel')
	response['Content-Disposition'] = 'attachment; filename="collectie.xlsx"'
	wb = Workbook()
	ws = wb.active
	ws.title = 'Info'
	ws['A1'] = "list of sheet tab codes"
	ws.append(['sheet number', 'quality check'])
	ws.append(['#001', 'instellingsnaam != YM, SM, MGB of OM'])
	ws.append(['#005', 'OBJECTNAAM'])
	ws.append(['005_01', 'objectnaam ontbreekt'])
	ws.append(['005_02', 'objectnaam start met hoofdletter'])
	ws.append(['005_03', 'objectnaam bevat lege occurences'])
	ws.append(['#006', 'TITEL'])
	ws.append(['#006_01', 'titel ontbreekt'])
	ws.append(['#006_02', 'foutieve start titel'])
	ws.append(['#006_03', 'titel eindigt op punt'])
	ws.column_dimensions['A'].width = 25
	ws.column_dimensions['B'].width = 60
	ym_001 = quality.ym_001()
	df_001 = ym_001
	ym_005 = quality.ym_005()
	df_005_01 = ym_005[0]
	df_005_02 = ym_005[1]
	df_005_03 = ym_005[2]
	ym_006 = quality.ym_006()
	df_006_01 = ym_006[0]
	df_006_02 = ym_006[1]
	df_006_03 = ym_006[2]
	if df_001.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#001")
		rows = dataframe_to_rows(df_001, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_005_01.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#005_01")
		rows = dataframe_to_rows(df_005_01, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_005_02.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#005_02")
		rows = dataframe_to_rows(df_005_02, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_005_03.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#005_03")
		rows = dataframe_to_rows(df_005_03, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_006_01.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#006_01")
		rows = dataframe_to_rows(df_006_01, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_006_02.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#006_02")
		rows = dataframe_to_rows(df_006_02, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_006_03.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#006_03")
		rows = dataframe_to_rows(df_006_03, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	wb.save(response)
	return response

def mmq001(request):
	response = HttpResponse(content_type='application/ms-excel')
	response['Content-Disposition'] = 'attachment; filename="collectie.xlsx"'
	wb = Workbook()
	ws = wb.active
	ws.title = 'Info'
	ws['A1'] = "list of sheet tab codes"
	ws.append(['sheet number', 'quality check'])
	ws.append(['#001', 'instellingsnaam != Hotel-Museum Arthur Merghelynck'])
	ws.append(['#005', 'OBJECTNAAM'])
	ws.append(['005_01', 'objectnaam ontbreekt'])
	ws.append(['005_02', 'objectnaam start met hoofdletter'])
	ws.append(['005_03', 'objectnaam bevat lege occurences'])
	ws.append(['#006', 'TITEL'])
	ws.append(['#006_01', 'titel ontbreekt'])
	ws.append(['#006_02', 'foutieve start titel'])
	ws.append(['#006_03', 'titel eindigt op punt'])
	ws.column_dimensions['A'].width = 25
	ws.column_dimensions['B'].width = 60
	mm_001 = quality.mm_001()
	df_001 = mm_001
	mm_005 = quality.mm_005()
	df_005_01 = mm_005[0]
	df_005_02 = mm_005[1]
	df_005_03 = mm_005[2]
	mm_006 = quality.ym_006()
	df_006_01 = mm_006[0]
	df_006_02 = mm_006[1]
	df_006_03 = mm_006[2]
	if df_001.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#001")
		rows = dataframe_to_rows(df_001, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_005_01.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#005_01")
		rows = dataframe_to_rows(df_005_01, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_005_02.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#005_02")
		rows = dataframe_to_rows(df_005_02, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_005_03.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#005_03")
		rows = dataframe_to_rows(df_005_03, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_006_01.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#006_01")
		rows = dataframe_to_rows(df_006_01, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_006_02.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#006_02")
		rows = dataframe_to_rows(df_006_02, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	if df_006_03.empty == True:
		print('empty dataframe')
	else:
		ws = wb.create_sheet("#006_03")
		rows = dataframe_to_rows(df_006_03, index=False)
		for r_idx, row in enumerate(rows, 1):
			for c_idx, value in enumerate(row, 1):
				ws.cell(row=r_idx, column=c_idx, value=value)
	wb.save(response)
	return response

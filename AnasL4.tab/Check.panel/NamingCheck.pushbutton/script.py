""" Verifica la correttezza della nomenclatura degli elementi di modello"""

__title__= 'Check\nNaming'
__author__= 'Roberto Dolfini'

import sys
import os
import csv
import clr
import System
import re # PER RIMUOVERE I CARATTERI ASCII NON VOLUTI

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
doc =  DocumentManager.Instance.CurrentDBDocument
 
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
 
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

from difflib import SequenceMatcher, Differ

import pyrevit
from pyrevit import script
from pyrevit import forms
import rpw.ui.forms
from rpw.ui.forms import (FlexForm, Label, Separator, Button, CheckBox)

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel

def similar(a, b):
	return SequenceMatcher(None, a, b).ratio()


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
activeView = doc.ActiveView
Output = script.get_output()


CategorieNaming = ["Telaio strutturale","Appoggi","Fondazioni strutturali","Muri","Pavimenti","Ringhiere","Pilastri strutturali","Tetti","Piloni","Modelli generici","Collegamenti strutturali","Armatura strutturale","Apparecchi elettrici","Cavi","Attrezzatura elettrica","Attrezzature speciali"]
CategorieNamingId = [-2001320,-2006138,-2001300,-2000011,-2000032,-2000126,-2001330,-2000035,-2006131,-2000151,-2009030,-2009000,-2001060,-2008039,-2001040,-2001350]
DizionarioBase = {}
for n, i in zip(CategorieNaming,CategorieNamingId):
	if n not in DizionarioBase:
		DizionarioBase[n] = i
	

## ANALISI DEL MODELLO
CollectorProgetto = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

CategorieInVista = []
DatiInVista = []
NomiInVista = []
LoadableInVista = []
ElementiInVista = []
error = []

for e in CollectorProgetto:
	temp = []
	Skipped = []
	if e != None and e.Category != None and e.Category.Name in CategorieNaming :
		if e.LookupParameter("Workset").AsValueString() != "03_Nascosto" and e.LookupParameter("Workset").AsValueString() != "02_Temporaneo":
			try:
				if e.Symbol:
					CategorieInVista.append(e.Category.Name)
					temp.append(e.LookupParameter("Famiglia").AsValueString())
					temp.append(e.Id)
					temp.append(True)
					DatiInVista.append(temp)
			except:
				try:
					CategorieInVista.append(e.Category.Name)
					temp.append(e.LookupParameter("Tipo").AsValueString())
					temp.append(e.Id)
					temp.append(False)
					DatiInVista.append(temp)
				except Exception as error:
					error.append(error)

for c in CategorieNaming:
	if c not in CategorieInVista:
		Skipped.append(c)


## ANALISI FOGLIO EXCEL

FileLocation = forms.pick_excel_file(save = False, title = "Seleziona Excel Naming")

ex = Excel.ApplicationClass()
WorkBook = ex.Workbooks.Open(FileLocation)
WorkSheet = WorkBook.Sheets("NOMENCLATURA")

CaselleUsate = WorkSheet.UsedRange
cell_value = CaselleUsate.Columns.Value2



Righe = []
try:
	for r in range(CaselleUsate.Rows.Count):
		temp_colonna = []
		for c in range(CaselleUsate.Columns.Count):
			
			if cell_value[r,c] == None or cell_value[r,c] == -2146826246 or cell_value[r,c] == -2146826259:
				pass
			else:
				temp_colonna.append(cell_value[r,c])
		if len(temp_colonna) <= 0:
			pass
		else:
			if str(temp_colonna[-1])[-1].isnumeric() or temp_colonna[-1].isupper():
				pass
			else:
				Righe.append(temp_colonna[-2:])
	WorkBook.Close()
except:
	WorkBook.Close()
	pass



Categoria = []
Naming = []
for riga in Righe:
	Categoria.append(riga[-1])
	Naming.append(riga[0])
DizionarioRiferimento = {}
for C,N  in zip(Categoria,Naming):
	if C not in DizionarioRiferimento.keys():
		DizionarioRiferimento[C] = [N]
	else:
		DizionarioRiferimento[C].append(N)

components = [CheckBox('NamingSuModello', 'Verifica nomenclatura elementi su modello'),CheckBox('NamingDaExcel', 'Verifica elementi su Excel mancanti nel modello'),Button('Conferma')]
form = FlexForm("Seleziona un'opzione",components)
form.show()

if form.values["NamingSuModello"]:
	
	# CONTROLLO EVENTUALI CATEGORIE NON PRESENTI SU EXCEL
	TotalmenteAssenti = []

	for c in CategorieInVista:
		if c not in DizionarioRiferimento:
			TotalmenteAssenti.append(c)

	TotalmenteAssenti = list(set(TotalmenteAssenti))
	DataAssenti = []

	for t in TotalmenteAssenti:
		ElementiMancanti = FilteredElementCollector(doc).OfCategoryId(ElementId(DizionarioBase[t])).WhereElementIsNotElementType().ToElements()
		for e in ElementiMancanti:
			try:
				e.Symbol
				Datatemp =[]
				Datatemp.append(e.Category.Name)
				Datatemp.append(e.LookupParameter("Famiglia").AsValueString())
				Datatemp.append("CATEGORIA ASSENTE SU EXCEL")

				DataAssenti.append(Datatemp)
			except:
				Datatemp =[]
				Datatemp.append(e.Category.Name)
				Datatemp.append(e.LookupParameter("Tipo").AsValueString())
				Datatemp.append("CATEGORIA ASSENTE SU EXCEL")

				DataAssenti.append(Datatemp)




	## COMPARAZIONE NOMENCLATURA & CERCA SIMILE

	ErrorElement = []
	ErrorCategory = []
	DataTable = []

	for lista in DatiInVista:	
		try:
			if lista[0] not in DizionarioRiferimento[doc.GetElement(lista[1]).Category.Name]:
				if lista[0] not in ErrorElement:
					ErrorElement.append(lista[0])
					ErrorCategory.append(doc.GetElement(lista[1]).Category.Name)
		except:
			pass

	error = 0
	for ctm in TotalmenteAssenti:
		if ctm not in ErrorCategory:
			error += 1
	if error != 0:
		Output.print_md("#:prohibited: ERRORE RILEVATO NELLA NOMENCLATURA")
		Output.print_md("###Elementi la cui categoria risulta essere assente su Excel")
		Output.print_table(
			table_data = DataAssenti,
			columns = ["CATEGORIA","NOME NEL MODELLO","PRESENTE SU EXCEL ?"],
			formats = ["","",""]
			)
		Output.insert_divider()
	
	### 

	for e,c in zip(ErrorElement,ErrorCategory):
		SaveSimilar = []
		DataTemp = []
		for Rif in DizionarioRiferimento[c]:
			SaveSimilar.append([similar(Rif,e),Rif])
		Valore = SaveSimilar[SaveSimilar.index(max(SaveSimilar))][0]
		Riferimento = SaveSimilar[SaveSimilar.index(max(SaveSimilar))][1]
		if Valore > 0.8:
			DataTemp.append(c)
			DataTemp.append(e)
			DataTemp.append(Riferimento)	
			DataTemp.append(str(round(Valore*100))+"%")	
		else:
			DataTemp.append(c)
			DataTemp.append(e)
			DataTemp.append("NON TROVATO IN EXCEL")
			DataTemp.append("N.D.")
		DataTable.append(DataTemp)

## GENERAZIONE TABELLA

	if DataTable:
		if error == 0:
			Output.print_md("#:prohibited: ERRORE RILEVATO NELLA NOMENCLATURA")
		Output.print_md("###Elementi nel modello, non presenti su Excel")
		Output.print_table(
				table_data = DataTable,
				columns = ["CATEGORIA","NOME NEL MODELLO","NOME DA EXCEL","PERCENTUALE SIMILARITA'"],
				formats = ["","","","{}"]
				)

		Output.insert_divider()
	else:
		Output.insert_divider()
		Output.print_md("#:white_heavy_check_mark: Non sono stati rilevati errori di nomenclatura")
		Output.insert_divider()

if form.values["NamingDaExcel"]: 
	DataEsclusi = []
	for c,c_id in zip(CategorieNaming,CategorieNamingId):
		FilterTemp = FilteredElementCollector(doc).OfCategoryId(ElementId(c_id)).WhereElementIsNotElementType().ToElements()
		NomiTemp=[]
		for e in FilterTemp:
			try:
				if e.Symbol:
					NomiTemp.append(e.LookupParameter("Famiglia").AsValueString())
			except:
					NomiTemp.append(e.LookupParameter("Tipo").AsValueString())
		EsclusiNomiTemp = []
		EsclusiCategoriaTemp = []
		try:
			temp = []
			for r in DizionarioRiferimento[c]:
				if r not in NomiTemp:
					temp.append(r)
				else:
					pass
			EsclusiCategoriaTemp.append(c)
			EsclusiNomiTemp.append(temp)
		except:
			pass
		temp = []
		for l1,l2 in zip(EsclusiCategoriaTemp,EsclusiNomiTemp):
			if l2:
				DataEsclusi.append([l1,l2])
	if DataEsclusi:
		Output.print_md("###Elementi presenti su Excel e non sul modello")
		Output.print_table(
				table_data = DataEsclusi,
				columns = ["CATEGORIA","NOME DA EXCEL"],
				)
	else:
		Output.insert_divider()
		Output.print_md("#:white_heavy_check_mark: Tutti gli elementi su Excel sono presenti nel modello")
		Output.insert_divider()
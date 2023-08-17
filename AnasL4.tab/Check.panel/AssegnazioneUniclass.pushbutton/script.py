""" Assegna il codice Uniclass al parametro di istanza e di tipo corrispondente"""
__title__= 'Assegnazione Uniclass\n L-4'
__author__= 'Roberto Dolfini'

# -*- coding: utf-8 -*-

import sys
import csv
import clr
import System
import re # PER RIMUOVERE I CARATTERI ASCII NON VOLUTI
import os
current_dir = os.path.dirname(__file__) # LOCATION ATTUALE DELLO SCRIPT

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType

clr.AddReference('RevitServices')
import RevitServices

from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *

import pyrevit
from pyrevit import script
from pyrevit import forms
import rpw.ui.forms
from rpw.ui.forms import (FlexForm, Label, Separator, Button, CheckBox)
## RAGGRUPPAMENTO DEFINIZIONI

def find(name, path):
	for root, dirs, files in os.walk(path):
		if name in files:
			return True
		else:
			return False
		
def estraiCodiceElemento(stringa):
	return stringa.split(".")[3][:3]

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
t_Transazione = Transaction(doc,"Assegnazione Uniclass")

if find("Uniclass.csv",current_dir):
	Codici = []
	Uniclass = []	
	with open(os.path.dirname(__file__)+"\\Uniclass.csv") as csvfile:
		reader = csv.reader(csvfile, delimiter=";")
		for row in reader:
			Uniclass.append(row[0])
			Codici.append(row[1])
	activeView = doc.ActiveView
	ElementiModello = FilteredElementCollector(doc,activeView.Id).WhereElementIsNotElementType().ToElements()
		
	isLoadable = []
	CodiceElemento = []
	ElementoScelto = []

	for elemento in ElementiModello:
		try:
			elemento.Symbol
			isLoadable.append(True)
		except:
			isLoadable.append(False)
	for bool,elemento in zip(isLoadable,ElementiModello):
		if bool:
			try:
				if "." in elemento.LookupParameter("Famiglia").AsValueString():
					CodiceElemento.append(estraiCodiceElemento(elemento.LookupParameter("Famiglia").AsValueString()))
					ElementoScelto.append(elemento)
			except:
				pass
		else:
			try:
				if "." in elemento.LookupParameter("Tipo").AsValueString():
					CodiceElemento.append(estraiCodiceElemento(elemento.LookupParameter("Tipo").AsValueString()))
					ElementoScelto.append(elemento)
			except:
				pass
	t_Transazione.Start()

	for elemento,codice in zip(ElementoScelto,CodiceElemento):	
		for sigla,uni in zip(Codici,Uniclass):
			if codice == sigla:
				doc.GetElement(elemento.GetTypeId()).LookupParameter("Codice assieme").Set(uni)
				elemento.LookupParameter("IDE_Uniclass").Set(uni)

	pyoutput = script.get_output()
	pyrevit.forms.toaster.send_toast("Codice Uniclass assegnato correttamente", title = "Assegna Uniclass", icon = sys.path[0] + "/iconanera.png")			
	t_Transazione.Commit()


		
else:
	#Alert('Generare CSV codici da Modello Dati', title="Database mancante", header="Uniclass.csv mancante")
	avviso = pyrevit.forms.alert(msg = "Generare un nuovo database da Excel ?\n L'operazione potrebbe richiedere fino a 2 minuti.", title = "Database mancante", warn_icon = True, exitscript = True, ok=True, cancel = True)


	if avviso:

		## GENERAZIONE DEL DATABASE PER VELOCIZZARE LE OPERAZIONI DI CONTROLLO

		clr.AddReference("Microsoft.Office.Interop.Excel")
		import Microsoft.Office.Interop.Excel as Excel
		FileLocation = forms.pick_excel_file(save = False, title = "Seleziona Modello Dati")
		try:
			
			ex = Excel.ApplicationClass()
			WorkBook = ex.Workbooks.Open(FileLocation)
			WorkSheet = WorkBook.Sheets("Constant+UniClass")
			ex.Visible = False
			Codici = []
			Parametri = []

			with forms.ProgressBar(indeterminate = True, title = "Salvataggio CSV") as pb:
				for row in range(2,83): # AGGIUNGERE QUI LE RIGHE DELL'EXCEL
					Parametri.append(WorkSheet.Rows[row].Value2[0,6])
					Codici.append(WorkSheet.Rows[row].Value2[0,5])
			WorkBook.Close()

		except Exception as Error:
			print(Error)
			WorkBook.Close()	

		# SALVATAGGIO CSV DI DATABASE
		Scrittura = []
		for c,p in zip(Codici,Parametri):
			Scrittura.append([c+";"+p])
		

		f = open(os.path.dirname(__file__)+"\\Uniclass.csv", mode = "w")
		writer = csv.writer(f)
		for riga in Scrittura:
			writer.writerow(riga)
		pyrevit.forms.toaster.send_toast("Database generato con successo", title = "Assegnazione Uniclass", icon = sys.path[0] + "/iconanera.png")
	else:
		pass

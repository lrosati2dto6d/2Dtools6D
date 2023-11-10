# -*- coding: utf-8 -*-

__title__= 'Check Compilazione\n L-4'
__author__= 'Roberto Dolfini'

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

def remove_non_ascii(text):
	return re.sub(r'[^\x00-\x7F]', ' ', text)

def find(name, path):
	for root, dirs, files in os.walk(path):
		if name in files:
			return True
		else:
			return False

def estraiParametri(elemento):

	nomi = [parametro.Definition.Name for parametro in elemento.ParametersMap if parametro.IsShared]
	guid = [parametro.GUID for parametro in elemento.ParametersMap if parametro.IsShared]
	return nomi,guid

def checkPresenzaParametro(listacheck,listarif):
	parametriDiTroppo = []
	for parametro in listacheck:
		if parametro not in listarif:
			parametriDiTroppo.append(parametro)
			
	return parametriDiTroppo

def estraiSigle(stringa):
	splittata = stringa.split(".")
	joinata = splittata[0]+"."+splittata[2]+"."+splittata[3][:3]
	return joinata

def uw(element):
	out = []
	if isinstance(element, list):
		for item in element:
			out.append(UnwrapElement(item))
		return out
	else:
		out=UnwrapElement(element)
		return out
		

def slice_parameters(counter, parameters):
	# Initialize the result list to store the sliced sublists
	result = []

	zero_indices = [i for i, val in enumerate(counter) if val == 0]


	for i in range(len(zero_indices)):
		if i == len(zero_indices) - 1:

			result.append(parameters[zero_indices[i]:])
		else:

			result.append(parameters[zero_indices[i]:zero_indices[i + 1]])

	return result

def GroupByKey(lista,key):

	savedKeys = []
	savedParams = []
	counterList = []
	for index,(k,l) in enumerate(zip(key,lista)):
		if k not in savedKeys:
			savedKeys.append(k)
			savedParams.append(l)
			counter = 0		
			counterList.append(counter)
		else:
			savedParams.append(l)
			counter += 1
			counterList.append(counter)
	RegroupList = []
		
	return savedKeys,slice_parameters(counterList, savedParams)

def checkHasValue(listaParametri,listaElementi):
	check = []
	checkElements = []
	failure = []
		
	for elemento,sublistParametri in zip(listaElementi,listaParametri):
		temp = []
		for parametro in sublistParametri:
		
			try:
				if elemento.LookupParameter(parametro).HasValue and elemento.LookupParameter(parametro).AsValueString() != None and elemento.LookupParameter(parametro).AsValueString() != "":
					temp.append(parametro)
			except:
				if doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).HasValue and doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).AsValueString() != None and doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).AsValueString() != "":
					temp.append(parametro)
				else:
					failure.append("ERRORE")
					
		check.append(temp)
		checkElements.append(elemento)
		
	if len(failure) == 0:
		failure = "NON CI SONO ERRORI"
	return check, checkElements

def checkCompilare(listaParametri,listaElementi):
	check = []
	checkElements = []
	failure = []
		
	for elemento,sublistParametri in zip(listaElementi,listaParametri):
		temp = []
		for parametro in sublistParametri:
		
			try:
				if elemento.LookupParameter(parametro).HasValue and "COMPILARE" in elemento.LookupParameter(parametro).AsValueString():
					temp.append(parametro)
			except:
				if doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).HasValue and "COMPILARE" in doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).AsValueString():
					temp.append(parametro)
				else:
					failure.append("ERRORE")
					
		check.append(temp)
		checkElements.append(elemento)
		
	if len(failure) == 0:
		failure = "NON CI SONO ERRORI"
	return check, checkElements
		
def checkZero(listaParametri,listaElementi):
	check = []
	checkElements = []
	failure = []
		
	for elemento,sublistParametri in zip(listaElementi,listaParametri):
		temp = []
		for parametro in sublistParametri:
			try:
				if elemento.LookupParameter(parametro).HasValue and elemento.LookupParameter(parametro).StorageType != StorageType.String:
					#if elemento.LookupParameter(parametro).AsValueString() == "0.00 m³" or elemento.LookupParameter(parametro).AsValueString() == "0.00 m²" or elemento.LookupParameter(parametro).AsValueString() == "0.00 m":
					if elemento.LookupParameter(parametro) != None:
						if str(elemento.LookupParameter(parametro).AsDouble()) == "0.0":
							temp.append(parametro)
			except:
				if doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).HasValue and doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).StorageType != StorageType.String:
					if doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro) != None:
						if str(doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).AsDouble()) == "0.0":
					#if doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).AsValueString() == "0.00 m³"or doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).AsValueString() == "0.00 m²" or doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).AsValueString() == "0.00 m" :
							temp.append(parametro)
				else:
					failure.append("ERRORE")
					
		check.append(temp)
		checkElements.append(elemento)
		
	if len(failure) == 0:
		failure = "NON CI SONO ERRORI"
	return check, checkElements

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


style = """
.legendgreen {
	color: Green; 
	font-size:15px;
}
"""
if find("Database.csv",current_dir):
	Codici = []
	Parametri = []	
	with open(os.path.dirname(__file__)+"\\Database.csv") as csvfile:
			reader = csv.reader(csvfile, delimiter=";")
			righe =[]
			for row in reader:
				righe.append(row)
			righe.sort(key=lambda riga: riga[0])
			for r in righe:
				Codici.append(r[0])
				Parametri.append(r[1])
	components = [Label('Predisposizione alla compilazione parametri:'),CheckBox('compilareCheck', 'Compilazione di default'),Separator(),Label('Opzioni riconoscimento parametri in eccesso:'),CheckBox('eccessoCheck', 'Verifica parametri in eccesso'),CheckBox('rimuoviEccesso', 'Rimozione stringhe in eccesso'),Button('Conferma')]
	form = FlexForm("Seleziona un'opzione",components)
	form.show()

	################################################################################################################
	activeView = doc.ActiveView
	ElementiModello = FilteredElementCollector(doc,activeView.Id).WhereElementIsNotElementType().ToElements()

	############## DIZIONARIO CODICI #####################################

	TotalRegroup = GroupByKey(Parametri,Codici)

	codice = TotalRegroup[0]
	gruppi = TotalRegroup[1]

	dizionarioCodici = {}
	for g,c in zip (gruppi,codice):
		dizionarioCodici[c] = g
		
	######################################################################

	############## PULIZIA ELEMENTI ######################################

	PickElements = []
	PickNames = []
	for elemento in ElementiModello:
		
		try:
			elemento.Symbol
			if "." in elemento.LookupParameter("Famiglia").AsValueString():
				PickElements.append(elemento)
				PickNames.append(elemento.LookupParameter("Famiglia").AsValueString())
		except:
			try:
				if "." in elemento.LookupParameter("Tipo").AsValueString():
					PickElements.append(elemento)
					PickNames.append(elemento.LookupParameter("Tipo").AsValueString())
			except:
				pass
	############## ESTRAZIONE SIGLE #####################################

	Caricabile = []

	for elemento in PickElements:
		try:
			elemento.Symbol
			Caricabile.append(True)
		except:
			Caricabile.append(False)

	Codice = []
	for bool,elemento in zip(Caricabile,PickElements):
		if bool:
			try:
				Codice.append(estraiSigle(elemento.LookupParameter("Famiglia").AsValueString()))
			except:
				pass
		else:
			try:
				Codice.append(estraiSigle(elemento.LookupParameter("Tipo").AsValueString()))
			except:
				pass

	######################################################################




#→	→	→	→	→	→	→	→	→	→	ASSEGNARE "COMPILARE" 

	failureCheck = []
	if form.values["compilareCheck"]: 
		t_Compilare = Transaction(doc,"Inserimento •••COMPILARE•••")
		t_Compilare.Start()
		#TransactionManager.Instance.EnsureInTransaction(doc)
				
		for elemento,codice in zip(PickElements,Codice):
			for parametro in dizionarioCodici[codice]:
				if elemento.LookupParameter(parametro):
					if elemento.LookupParameter(parametro).Definition.Name == "TEC_Numero seriale":
						elemento.LookupParameter(parametro).Set("-")
					elif elemento.LookupParameter(parametro).AsValueString() == None or elemento.LookupParameter(parametro).AsValueString() == " " or elemento.LookupParameter(parametro).AsValueString() == "":
						elemento.LookupParameter(parametro).Set("•••COMPILARE•••")

				else:
					try:
						if doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).StorageType == StorageType.String and doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).AsValueString() == None or doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).AsValueString() == " " or doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).AsValueString() == "":
							doc.GetElement(elemento.GetTypeId()).LookupParameter(parametro).Set("•••COMPILARE•••")
					except:
						failureCheck.append(elemento)
						failureCheck.append(parametro)
		
		#TransactionManager.Instance.TransactionTaskDone()

		
		NAMEParameterMap = (estraiParametri(elemento)[0] for elemento in PickElements)
		GUIDParameterMap = (estraiParametri(elemento)[1] for elemento in PickElements)
		
		parametroDiTroppo = []
		saveElement = []
		for codice,elemento,listaparametri in zip(Codice,PickElements,NAMEParameterMap):
			temp = []
			for parametro in listaparametri:
				if parametro not in dizionarioCodici[codice]:
					temp.append(parametro)
			if len(temp) > 0:
				parametroDiTroppo.append(temp)
				saveElement.append(elemento)
		
		valorizzatiDiTroppo = []
		elementiDiTroppo = []

		for sublist,elemento in zip(parametroDiTroppo,saveElement):
			temp = []
			for parametro in sublist:
				if  not elemento.LookupParameter(parametro).HasValue or elemento.LookupParameter(parametro).AsValueString() == None or elemento.LookupParameter(parametro).AsValueString() == "":
					pass
				else:
					temp.append(parametro)

			if len(temp) > 0:
				elementiDiTroppo.append(elemento)
				valorizzatiDiTroppo.append(temp)

		# ASSEGNARE VALORI DI TIPO

		Opera = []
		Parteopera = []
		Elemento = []
		Struttura = []
		Result = []

		for name in PickNames:
			Opera.append(name.split(".")[0])
			Parteopera.append(name.split(".")[2])
			Elemento.append(name.split(".")[3][:3])
			Struttura.append(name.split(".")[1])

		try:
			for elemento,o,po,e,s in zip(PickElements,Opera,Parteopera,Elemento,Struttura):
				doc.GetElement(elemento.GetTypeId()).LookupParameter("Modello").Set(o)
				doc.GetElement(elemento.GetTypeId()).LookupParameter("Commenti sul tipo").Set(po)
				doc.GetElement(elemento.GetTypeId()).LookupParameter("Descrizione").Set(e)
				if s != "XXX":
					doc.GetElement(elemento.GetTypeId()).LookupParameter("Contrassegno tipo").Set(s)
				else:
					doc.GetElement(elemento.GetTypeId()).LookupParameter("Contrassegno tipo").Set("")
			Result.append("Success")
		except:
			Result.append("Failure")

		t_Compilare.Commit()
		pyrevit.forms.toaster.send_toast("Compilazione Effettuata", title = "Check Compilazione", icon = sys.path[0] + "/iconanera.png")
############################################################# CHECK DEI PARAMETRI IN ECCESSO

	if form.values["eccessoCheck"]:
		
		NAMEParameterMap = (estraiParametri(elemento)[0] for elemento in PickElements)
		GUIDParameterMap = (estraiParametri(elemento)[1] for elemento in PickElements)
		
		parametroDiTroppo = []
		saveElement = []
		for codice,elemento,listaparametri in zip(Codice,PickElements,NAMEParameterMap):
			temp = []
			for parametro in listaparametri:
				if parametro not in dizionarioCodici[codice]:
					temp.append(parametro)
			if len(temp) > 0:
				parametroDiTroppo.append(temp)
				saveElement.append(elemento)
		
		valorizzatiDiTroppo = []
		elementiDiTroppo = []

		for sublist,elemento in zip(parametroDiTroppo,saveElement):
			temp = []
			for parametro in sublist:
				if  not elemento.LookupParameter(parametro).HasValue or elemento.LookupParameter(parametro).AsValueString() == None or elemento.LookupParameter(parametro).AsValueString() == "":
					pass
				else:
					temp.append(parametro)

			if len(temp) > 0:
				elementiDiTroppo.append(elemento)
				valorizzatiDiTroppo.append(temp)

		output=zip(*checkHasValue(valorizzatiDiTroppo,elementiDiTroppo)) # CHECK DI TROPPO
		cleanDiTroppo = []
		for list in output:
			if len(list[0]) == 0:
				pass
			else:
				cleanDiTroppo.append(list)

		nomiParametri = []
		for elemento in PickElements:
			try:
				nomiParametri.append(estraiParametri(elemento)[0])
			except:
				pass

		output=zip(*checkCompilare(nomiParametri,PickElements)) # CHECK COMPILARE DA RIMUOVERE
		cleanCompilare = []
		for list in output:
			if len(list[0]) == 0:
				pass
			else:
				cleanCompilare.append(list)
		
		output=zip(*checkZero(nomiParametri,PickElements)) # CHECK ZERI DA RIMUOVERE
		cleanZeri = []
		for list in output:
			if len(list[0]) == 0:
				pass
			else:
				cleanZeri.append(list)
########################################### OUTPUT DI PARAMETRI IN ECCESSO ################################################################		
		pyoutput = script.get_output()
		if len(cleanDiTroppo) != 0:
			pyoutput.add_style(style)
			pyoutput.print_md("#:prohibited: PARAMETRI IN ECCESSO :")
			pyoutput.insert_divider()
			for lista in cleanDiTroppo:
				listaparametri = [remove_non_ascii(x) for x in lista[0]]
				pyoutput.print_md("**{} - {} : {}**".format(doc.GetElement(lista[1].Id).LookupParameter("Famiglia").AsValueString(),doc.GetElement(lista[1].Id).LookupParameter("Tipo").AsValueString(),pyoutput.linkify(lista[1].Id)))
				for x in listaparametri:
					pyoutput.print_md("**• {}**".format(x))
				pyoutput.insert_divider()
		else:
			pyoutput.insert_divider()
			pyoutput.print_html('<div class="legendgreen"><b>Non sono stati trovati parametri in eccesso.</b></div>')
			pyoutput.insert_divider()
########################################### OUTPUT DI COMPILARE IN ECCESSO #############################################################
		pyoutput = script.get_output()
		if len(cleanCompilare) != 0:
			pyoutput.add_style(style)
			pyoutput.print_md("#:prohibited: COMPILARE IN ECCESSO :")
			pyoutput.insert_divider()
			for lista in cleanCompilare:
				listaparametri = [remove_non_ascii(x) for x in lista[0]]
				pyoutput.print_md("**{} - {} : {}**".format(doc.GetElement(lista[1].Id).LookupParameter("Famiglia").AsValueString(),doc.GetElement(lista[1].Id).LookupParameter("Tipo").AsValueString(),pyoutput.linkify(lista[1].Id)))
				for x in listaparametri:
					pyoutput.print_md("**• {}**".format(x))
				pyoutput.insert_divider()
		else:
			pyoutput.insert_divider()
			pyoutput.print_html('<div class="legendgreen"><b>Non sono stati trovati "COMPILARE" in eccesso.</b></div>')
			pyoutput.insert_divider()
########################################### OUTPUT DI ZERI IN ECCESSO #############################################################
		pyoutput = script.get_output()
		if len(cleanZeri) != 0:
			pyoutput.add_style(style)
			pyoutput.print_md("#:prohibited: ZERI IN ECCESSO :")
			pyoutput.insert_divider()

			for lista in cleanZeri:
				listaparametri = [remove_non_ascii(x) for x in lista[0]]
				pyoutput.print_md("**{} - {} : {}**".format(doc.GetElement(lista[1].Id).LookupParameter("Famiglia").AsValueString(),doc.GetElement(lista[1].Id).LookupParameter("Tipo").AsValueString(),pyoutput.linkify(lista[1].Id)))
			
				for x in listaparametri:
					pyoutput.print_md("**• {}**".format(x))
				pyoutput.insert_divider()
		else:
			pyoutput.insert_divider()
			pyoutput.print_html('<div class="legendgreen"><b>Non sono stati trovati "0" in eccesso.</b></div>')
			pyoutput.insert_divider()		
		if len(cleanDiTroppo) == 0 and len(cleanCompilare) == 0 and len(cleanZeri) == 0:
			pyoutput.print_md("#:white_heavy_check_mark: NON SONO STATI RISCONTRATI ERRORI NELLA COMPILAZIONE PARAMETRI")


	if form.values["rimuoviEccesso"]:
		
		NAMEParameterMap = (estraiParametri(elemento)[0] for elemento in PickElements)
		#GUIDParameterMap = (estraiParametri(elemento)[1] for elemento in PickElements)
		

		parametroDiTroppo = []
		saveElement = []
		
		for codice,elemento,listaparametri in zip(Codice,PickElements,NAMEParameterMap):
			temp = []
			for parametro in listaparametri:
				if parametro not in dizionarioCodici[codice]:
					temp.append(parametro)

			if len(temp) > 0:
				parametroDiTroppo.append(temp)
				saveElement.append(elemento)
		
		valorizzatiDiTroppo = []
		elementiDiTroppo = []

		for sublist,elemento in zip(parametroDiTroppo,saveElement):
			temp = []
			
			for parametro in sublist:
				if  not elemento.LookupParameter(parametro).HasValue or elemento.LookupParameter(parametro).AsValueString() == None or elemento.LookupParameter(parametro).AsValueString() == "":
					pass
				else:
					temp.append(parametro)

			if len(temp) > 0:
				elementiDiTroppo.append(elemento)
				valorizzatiDiTroppo.append(temp)

		t_Rimozione = Transaction(doc,"Rimozione parametri in eccesso")
		
		NAMEParameterMap = (estraiParametri(elemento)[0] for elemento in PickElements)
		#GUIDParameterMap = (estraiParametri(elemento)[1] for elemento in PickElements)

		t_Rimozione.Start()

		nomiParametri = []
		for elemento in PickElements:
			try:
				nomiParametri.append(estraiParametri(elemento)[0])
			except:
				pass

		output=zip(*checkCompilare(nomiParametri,PickElements))
		cleanCompilare = []
		
		for list in output:
			if len(list[0]) == 0:
				pass
			else:
				cleanCompilare.append(list)

		for list in cleanCompilare:
			for p in list[0]:
				list[1].LookupParameter(p).Set("")
		try:
			for elemento in PickElements: #VERIFICA PRESENZA IN PARAMETRI DI TIPO
				ListaParametriTipo = estraiParametri(doc.GetElement(elemento.GetTypeId()))[0]
				for parametro in ListaParametriTipo:
					print(elemento.LookupParameter(parametro).AsValueString())


		except Exception as e:
			pass

		for sublist,elemento in zip(valorizzatiDiTroppo,elementiDiTroppo):
			for parametro in sublist:
				elemento.LookupParameter(parametro).Set("")
				#pyoutput.print_md("#:white_heavy_check_mark: PARAMETRI IN ECCESSO RIMOSSI")
				
				pyoutput = script.get_output()
				pyrevit.forms.toaster.send_toast("Parametri in eccesso rimossi.", title = "Pulizia parametri", icon = sys.path[0] + "/iconanera.png")
		t_Rimozione.Commit()
		###		
		t = Transaction(doc,"Togli Compilare da Tipo")
		
		t.Start()
		for elemento in PickElements:
			#parametro.Definition.Name for parametro in elemento.ParametersMap if parametro.IsShared

					
			TypeParams = doc.GetElement(elemento.GetTypeId()).ParametersMap

			for tp in TypeParams:
				if tp.IsShared :
					try:
						if "COMPILARE" in tp.AsValueString():
							tp.Set("")
					except Exception as i:
						pass
						#print(i)
						try:
							if "COMPILARE" in tp.AsValueString():
								tp.Set("")
						except Exception as e:
							pass
							#print(e)
			
		t.Commit()

		###



		
		pyoutput = script.get_output()
		pyrevit.forms.toaster.send_toast("Parametri in eccesso rimossi.", title = "Pulizia parametri", icon = sys.path[0] + "/iconanera.png")
		
else:
	#Alert('Generare CSV codici da Modello Dati', title="Database mancante", header="Database.csv mancante")
	avviso = pyrevit.forms.alert(msg = "Generare un nuovo database da Excel ?\n L'operazione può richiedere fino a 5 minuti.", title = "Database mancante", warn_icon = True, exitscript = True, ok=True, cancel = True)


	if avviso:

		## GENERAZIONE DEL DATABASE PER VELOCIZZARE LE OPERAZIONI DI CONTROLLO

		clr.AddReference("Microsoft.Office.Interop.Excel")
		import Microsoft.Office.Interop.Excel as Excel
		FileLocation = forms.pick_excel_file(save = False, title = "Seleziona Modello Dati")
		try:
			
			#FileLocation = r"C:\Users\rdolfini\Documents\002.ANAS\001.Modello Dati\Allegato A_MODELLO DATI_2Dto6D.xlsx"
			ex = Excel.ApplicationClass()
			WorkBook = ex.Workbooks.Open(FileLocation)
			#start = timer()
			WorkSheet = WorkBook.Sheets("MODELLO DATI BIM")
			ex.Visible = False
			Codici = []
			Parametri = []

			with forms.ProgressBar(indeterminate = True, title = "Salvataggio CSV") as pb:
				for row in range(3,4348): # AGGIUNGERE QUI LE RIGHE DELL'EXCEL
					tempCode = []
					tempCode.append(WorkSheet.Rows[row].Value2[0,1])
					tempCode.append(WorkSheet.Rows[row].Value2[0,3])
					tempCode.append(WorkSheet.Rows[row].Value2[0,5])
					Parametri.append(remove_non_ascii(WorkSheet.Rows[row].Value2[0,17]))
					Codici.append(".".join(tempCode))
			WorkBook.Close()

		except Exception as Error:
			print(Error)
			WorkBook.Close()	

		# SALVATAGGIO CSV DI DATABASE
		Scrittura = []
		for c,p in zip(Codici,Parametri):
			Scrittura.append([c+";"+p])
		

		f = open(os.path.dirname(__file__)+"\\Database.csv", mode = "w")
		writer = csv.writer(f)
		for riga in Scrittura:
			writer.writerow(riga)
	else:
		pass
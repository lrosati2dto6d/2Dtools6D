""" Valorizzazione manuale dei campi costituenti la BMS """

__title__= 'Valorizzazione Manuale\nBMS'
__author__= 'Roberto Dolfini'

import codecs
import re
import unicodedata
from pyrevit import forms
import clr
import System

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

trans = Transaction(doc,"Assegnazione BMS Manuale")

selezione = uidoc.Selection.GetElementIds()

if len(selezione) < 1:
	forms.alert(msg = "Nessuna selezione attiva",sub_msg =  "Mantenere in selezione attiva gli elementi da valorizzare.", ok=True, exitscript = True)

modalita_di_inserimento = forms.ask_for_one_item(["Totalmente Manuale","Concatenazione Parametri Esistenti"], default="Totalmente Manuale", prompt="Seleziona come inserire i valori, totalmente manuale oppure \ncome concatenazione dei parametri pre-esistenti", title="Valorizzazione BMS", width = 1200)


if modalita_di_inserimento == "Totalmente Manuale":

	progressivoCampata = forms.ask_for_string(prompt = "Inserisci progressivo Campata", title="Progressivo CAM")
	progressivoImpalcato = forms.ask_for_string(prompt = "Inserisci progressivo Impalcato", title="Progressivo IMP")
	progressivoStruttura = forms.ask_for_string(prompt = "Inserisci progressivo Struttura Campata", title="Progressivo STC")

	elementi_da_valorizzare = []
	for s in selezione:
		try:
			doc.GetElement(s).Symbol
			if "." in doc.GetElement(s).LookupParameter("Famiglia").AsValueString() and "XXX" not in doc.GetElement(s).LookupParameter("Famiglia").AsValueString():
				elementi_da_valorizzare.append(s)

		except:
			if "." in doc.GetElement(s).LookupParameter("Tipo").AsValueString() and "XXX" not in doc.GetElement(s).LookupParameter("Tipo").AsValueString():
				elementi_da_valorizzare.append(s)			

	trans.Start()

	for s in elementi_da_valorizzare:
		progressivoElemento = doc.GetElement(s).LookupParameter("INF_Codice BMS").AsValueString().split(".")[-1]
		codiceElemento = doc.GetElement(doc.GetElement(s).GetTypeId()).LookupParameter("Descrizione").AsValueString()
		doc.GetElement(s).LookupParameter("INF_Codice BMS").Set("CAM."+progressivoCampata+".STC."+progressivoStruttura+"."+ codiceElemento +"."+progressivoElemento)
		doc.GetElement(s).LookupParameter("INF_Campata di appartenenza").Set(progressivoCampata)
		doc.GetElement(s).LookupParameter("INF_Impalcato di appartenenza").Set(progressivoImpalcato)
		doc.GetElement(s).LookupParameter("INF_Numerazione struttura campata").Set(progressivoStruttura)

	trans.Commit()


elif modalita_di_inserimento == "Concatenazione Parametri Esistenti":

	elementi_da_valorizzare = []
	for s in selezione:
		try:
			doc.GetElement(s).Symbol
			if "." in doc.GetElement(s).LookupParameter("Famiglia").AsValueString() and "XXX" not in doc.GetElement(s).LookupParameter("Famiglia").AsValueString():
				elementi_da_valorizzare.append(s)

		except:
			if "." in doc.GetElement(s).LookupParameter("Tipo").AsValueString() and "XXX" not in doc.GetElement(s).LookupParameter("Tipo").AsValueString():
				elementi_da_valorizzare.append(s)
	
	trans.Start()

	for s in elementi_da_valorizzare:
		progressivoCampata = doc.GetElement(s).LookupParameter("INF_Campata di appartenenza").AsValueString()
		progressivoStruttura = doc.GetElement(s).LookupParameter("INF_Numerazione struttura campata").AsValueString()
		progressivoElemento = doc.GetElement(s).LookupParameter("INF_Codice BMS").AsValueString().split(".")[-1]
		codiceElemento = doc.GetElement(doc.GetElement(s).GetTypeId()).LookupParameter("Descrizione").AsValueString()
		doc.GetElement(s).LookupParameter("INF_Codice BMS").Set("CAM." + progressivoCampata + ".STC." + progressivoStruttura + "." + codiceElemento + "." + progressivoElemento)

	trans.Commit()
else:
	pass



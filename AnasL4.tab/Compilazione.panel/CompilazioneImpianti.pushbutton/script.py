""" Compila parametri di default degli elementi appartenenti al workset Impianti """

__title__= 'Valorizzazione Default\nImpianti'
__author__= 'Roberto Dolfini'

import sys
import os
import csv
import clr
import System

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



import pyrevit
from pyrevit import script
from pyrevit import forms
#import rpw.ui.forms
#from rpw.ui.forms import (FlexForm, Label, Separator, Button, CheckBox)



def EstraiCodici(elemento):
	name = elemento.LookupParameter("Famiglia").AsValueString()

	Opera = name.split(".")[0]
	Parteopera = name.split(".")[2]
	Elemento = name.split(".")[3][:3]
	Struttura = name.split(".")[1]

	return Opera,Parteopera,Elemento,Struttura

def EstraiNome(elemento):
	try:
		elemento.Symbol
		return elemento.LookupParameter("Famiglia").AsValueString()
	except:
		return elemento.LookupParameter("Tipo").AsValueString()
	
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
activeView = doc.ActiveView
Output = script.get_output()

t_AssegnazioneParametriIstanza = Transaction(doc,"Assegnazione parametri impianti di istanza")
t_AssegnazioneParametriTipo = Transaction(doc,"Assegnazione parametri impianti di tipo")

## ESTRAZIONE ELEMENTI SU WORKSET IMPIANTI


CollectorWorkset = FilteredWorksetCollector(doc).ToWorksets()

for w in CollectorWorkset:
	if w.Name == "08_IMP":
		wksetId = w.Id		

filter = ElementWorksetFilter(wksetId)

ImpiantiInViewo = FilteredElementCollector(doc,activeView.Id).WhereElementIsNotElementType().WherePasses(filter).ToElements()

ImpiantiInView = []

for imp in ImpiantiInViewo:
	if imp.Category.Name != "Sistema di tubazioni":
		ImpiantiInView.append(imp)


## LISTE DI RIFERIMENTO

GEO_Quota_sensore = "ACC,IFS,SCA,SEM,TEC,TCM,VCM,TSS,TIG,SUM,CLI,BRE,IDO"
ANA_Progettista = "ALI,BLI,CAB,CNP,CAE,CAV,COL,COM,CEE,DIE,IMT,INS,INM,MUL,PLI,PZE,QEB,QEM,REL,REP,RIF,SCS,SDE,TRS"
TEC_Numero_seriale = "IFS,TEC,TCM,ACC,SCA,VCM,TSS,SEM,TIG,SUM,CLI,BRE,IDO"
TEC_Posizione = "ACC,IFS,SCA,SEM,TEC,TCM,VCM,TIG,SUM"
ElementiBMS = "AMS,AAP,BAG,BIN,CAS,CNT,COR,CUN,ISA,LOR,MUS,PPZ,PEN,PUL,PUN,RIS,SBL,SGE,SOL,SSB,SAR,SEL,TAN,TRV,TRA,VEL"
IDE_ElementoDiAppartenenza = "SGE,ACC,IFS,SCA,SEM,TEC,TCM,VCM,TSS,TIG,SUM,CLI,BRE,IDO"
TEC_Utilizzo = "CAE"


## ASSEGNAZIONE PARAMETRI DI ISTANZA

t_AssegnazioneParametriIstanza.Start()

CodiceElemento = []
for elemento in ImpiantiInView:
	CodiceElemento.append(doc.GetElement(elemento.GetTypeId()).LookupParameter("Descrizione").AsValueString())

for elemento,codice in zip(ImpiantiInView,CodiceElemento):

	elemento.LookupParameter("IDE_LOR").Set("BASSO") # Valore di default da assegnare a tutti
	
	if codice in GEO_Quota_sensore:
		elemento.LookupParameter("GEO_Quota sensore").Set(elemento.LookupParameter("Quota altimetrica da livello").AsDouble())

	if codice in ANA_Progettista:
		elemento.LookupParameter("ANA_Progettista").Set("RINA Consulting SpA")

	if codice in TEC_Numero_seriale:
		elemento.LookupParameter("TEC_Numero seriale").Set("-")

	if codice in TEC_Posizione:
		host = elemento.Host
		
		if EstraiNome(host).split(".")[1] != "XXX":
			elemento.LookupParameter("TEC_Posizione").Set("CAM." + host.LookupParameter("INF_Campata di appartenenza").AsString())
		else:
			elemento.LookupParameter("TEC_Posizione").Set(host.LookupParameter("IDE_Gruppo anagrafica").AsString()[6:12])
	
	if codice in IDE_ElementoDiAppartenenza:
		host = elemento.Host
		elemento.LookupParameter("IDE_Elemento di appartenenza").Set(doc.GetElement(host.GetTypeId()).LookupParameter("Descrizione").AsValueString()+"."+host.LookupParameter("IDE_Codice WBS").AsValueString().split(".")[-1])
		
	if codice == "CAE":
		elemento.LookupParameter("TEC_Utilizzo").Set(elemento.LookupParameter("Tipo di sistema").AsValueString())

t_AssegnazioneParametriIstanza.Commit()
"""
## ASSEGNAZIONE PARAMETRI DI TIPO

t_AssegnazioneParametriTipo.Start()

for elemento in ImpiantiInView:

	Codici = EstraiCodici(elemento)

	doc.GetElement(elemento.GetTypeId()).LookupParameter("Modello").Set(Codici[0])
	doc.GetElement(elemento.GetTypeId()).LookupParameter("Commenti sul tipo").Set(Codici[1])
	doc.GetElement(elemento.GetTypeId()).LookupParameter("Descrizione").Set(Codici[2])

	if "XXX" not in elemento.LookupParameter("Famiglia").AsValueString():
		doc.GetElement(elemento.GetTypeId()).LookupParameter("Contrassegno tipo").Set(Codici[3])
	else:
		doc.GetElement(elemento.GetTypeId()).LookupParameter("Contrassegno tipo").Set("")

t_AssegnazioneParametriTipo.Commit()
"""
pyrevit.forms.toaster.send_toast("Compilazione effettuata", title = "Compilazione Default Impianti", icon = sys.path[0] + "/iconanera.png")

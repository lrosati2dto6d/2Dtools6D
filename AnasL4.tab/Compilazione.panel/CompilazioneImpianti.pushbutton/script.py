""" Compila parametri di default degli elementi appartenenti al workset Impianti """

__title__= 'Valorizzazione Default\nImpianti'
__author__= 'Roberto Dolfini-Luca Rosati'

import sys
import clr

clr.AddReference("RevitServices")
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
	
doc = __revit__.ActiveUIDocument.Document # type: ignore
uidoc = __revit__.ActiveUIDocument # type: ignore
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

ImpiantiInViewo = FilteredElementCollector(doc,activeView.Id).WhereElementIsNotElementType().WherePasses(filter).ToElements() # type: ignore



ImpiantiInView = []
impno = []
for imp in ImpiantiInViewo:
	try:
		if imp.Category.Name != "Sistema di tubazioni" and imp.Category.Name != "Linea d'asse" :
			ImpiantiInView.append(imp)
	except:
		impno.append(imp)


## LISTE DI RIFERIMENTO

GEO_Quota_sensore = "ACC,IFS,SCA,SEM,TEC,TCM,VCM,TSS,TIG,SUM,CLI,BRE,IDO"
ANA_Progettista = "ALI,BLI,CAB,CNP,CAE,CAV,COL,COM,CEE,DIE,IMT,INS,INM,MUL,PLI,PZE,QEB,QEM,REL,REP,RIF,SCS,SDE,TRS"
TEC_Numero_seriale = "IFS,TEC,TCM,ACC,SCA,VCM,TSS,SEM,TIG,SUM,CLI,BRE,IDO"
TEC_Posizione = "ACC,IFS,SCA,SEM,TEC,TCM,VCM,TIG,SUM"
ElementiBMS = "AMS,AAP,BAG,BIN,CAS,CNT,COR,CUN,ISA,LOR,MUS,PPZ,PEN,PUL,PUN,RIS,SBL,SGE,SOL,SSB,SAR,SEL,TAN,TRV,TRA,VEL"
IDE_ElementoDiAppartenenza = "SGE,ACC,IFS,SCA,SEM,TEC,TCM,VCM,TSS,TIG,SUM,CLI,BRE,IDO"
TEC_Utilizzo = "CAE"


## ASSEGNAZIONE PARAMETRI DI ISTANZA
result = []
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
		
		try:
			if EstraiNome(host).split(".")[1] != "XXX":
				elemento.LookupParameter("TEC_Posizione").Set("CAM." + host.LookupParameter("INF_Campata di appartenenza").AsString())
			else:
				elemento.LookupParameter("TEC_Posizione").Set(host.LookupParameter("IDE_Gruppo anagrafica").AsString()[6:12])
		except:
			pass
	
	if codice in IDE_ElementoDiAppartenenza:
		try:
			host = elemento.Host
			elemento.LookupParameter("IDE_Elemento di appartenenza").Set(doc.GetElement(host.GetTypeId()).LookupParameter("Descrizione").AsValueString()+"."+host.LookupParameter("IDE_Codice WBS").AsValueString().split(".")[-1])
		except:
			result.append(elemento.Id)

	if codice == "CAE":
		type_el = doc.GetElement(elemento.GetTypeId())
		valore = type_el.LookupParameter("Contrassegno tipo").AsValueString()
		elemento.LookupParameter("TEC_Utilizzo").Set(valore)

t_AssegnazioneParametriIstanza.Commit()

output = script.get_output()
output.set_height(600)

if len(result) != 0:
	output.print_md("##\tTransfer Not Completed for These Elements:")
	for r in result:
		print(r)

if len(impno) != 0:
	output.print_md("##\tQuesti Elementi sono Errati:")
	for i in impno:
		print("{}".format(output.linkify(i.Id)))

"""Verifica della corretta numerazione degli elementi, BMS o Gruppo Anagrafica"""

__title__= 'Check Numerazione'
__author__= 'Roberto Dolfini'

import clr
import os

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.Creation import *

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

## RAGGRUPPAMENTO DEFINIZIONI

def PuntoMediano(punto0,punto1):
    return XYZ((punto0.X+punto1.X)/2,(punto0.Y+punto1.Y)/2,(punto0.Z+punto1.Z)/2)


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

t = Transaction(doc,"Check Numerazione")

opz = ["INF_Codice BMS","IDE_Gruppo anagrafica"]

sel = forms.ask_for_one_item(opz, default = opz[0], button_name = "Seleziona")

# INSERIMENTO TESTO

ActiveSelection = uidoc.Selection.GetElementIds()

text_note_type = FilteredElementCollector(doc).OfClass(TextNoteType).FirstElement()

t.Start()

for e in ActiveSelection:
    
    if doc.GetElement(e).LookupParameter(sel):
        param = doc.GetElement(e).LookupParameter(sel).AsString()[-3:]

    else:
        param = "N.D."
    
    try:
        TextNote.Create(doc, doc.ActiveView.Id, doc.GetElement(e).Location.Point, param, text_note_type.Id)
    
    except:
        TextNote.Create(doc, doc.ActiveView.Id, PuntoMediano(doc.GetElement(e).get_BoundingBox(None).Min,doc.GetElement(e).get_BoundingBox(None).Max), param, text_note_type.Id)

t.Commit()




"""
# METODOLOGIA TRAMITE COLORE

 

#GradienteRGB = [Color(255,0,0),Color(255,100,0),Color(255,200,0),Color(150,255,0),Color(0,255,205),Color(0,175,255),Color(0,0,255),Color(100,0,255),Color(255,0,200)]

GradienteRGB = [
Color(255, 0, 0),
Color(255, 0, 41),
Color(255, 0, 69),
Color(255, 0, 97),
Color(250, 0, 128),
Color(230, 0, 161),
Color(197, 0, 195),
Color(146, 0, 227),
Color(0, 0, 255)
]

Retino = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()

OverrideBase = OverrideGraphicSettings()
OverrideBase.SetSurfaceForegroundPatternId(Retino[0].Id)

Collector = FilteredElementCollector(doc,doc.ActiveView.Id).WhereElementIsNotElementType().ToElements()
ConvertiLista = list(Collector)

Numerazione = []
for elemento in Collector:
    try:
        temp = []
        temp.append(elemento)
        temp.append(elemento.LookupParameter("IDE_Gruppo anagrafica").AsValueString()[-3:])
        Numerazione.append(temp)
    except:
        pass
Numerazione.sort(key = Second)

Repeat = GradienteRGB * ((len(Collector)/len(GradienteRGB)) + len(GradienteRGB)) # ESTENDO LA LISTA DEI COLORI PER AVERNE ABBASTANZA 

t.Start()

for lista in Numerazione:
    OverrideBase.SetSurfaceForegroundPatternColor(Repeat[int(lista[1])])
    doc.ActiveView.SetElementOverrides(lista[0].Id,OverrideBase)

t.Commit() 
"""

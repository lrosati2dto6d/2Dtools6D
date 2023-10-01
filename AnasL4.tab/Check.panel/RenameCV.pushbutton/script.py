"""Rinomina Nomi e Parametri da CV a PV"""

__title__= 'Rename CV to PV'
__author__= 'Luca Rosati'

import clr
import System

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import *

clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(Revit.Elements)

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *

from collections import defaultdict
from pyrevit import HOST_APP
from pyrevit.framework import List
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script
from pyrevit.framework import Emojis

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument

rapp = doc.Application

#-------------------------------------DEFINIZIONI

def ExitScript(check):
	if not check:
		script.exit()

def ConvUnitsFM(number): #Feet to m
	output = number/0.3048
	return output

def ParaBuilt (element,builtinparameter):
	parameter = element.get_Parameter(BuiltInParameter.builtinparameter)
	return parameter

def Para(element,paraname):
	parameter = element.LookupParameter(paraname)
	return parameter

def ParaInst(element,paraname):
	if element.LookupParameter(paraname).StorageType == StorageType.Double:
		value = element.LookupParameter(paraname).AsDouble()
	elif element.LookupParameter(paraname).StorageType == StorageType.ElementId:
		value = element.LookupParameter(paraname).AsElementId()
	elif element.LookupParameter(paraname).StorageType == StorageType.String:
		value = element.LookupParameter(paraname).AsString()
	elif element.LookupParameter(paraname).StorageType == StorageType.Integer:
		value = element.LookupParameter(paraname).AsInteger()
	elif element.LookupParameter(paraname).StorageType == None:
		value = "Da Compilare"
	return value

def ParaType(element,paraname,document):
	element_type = document.GetElement(element.GetTypeId())
	if element_type.LookupParameter(paraname).StorageType == StorageType.Double:
		value = element_type.LookupParameter(paraname).AsDouble()
	elif element_type.LookupParameter(paraname).StorageType == StorageType.ElementId:
		value = element_type.LookupParameter(paraname).AsElementId()
	elif element_type.LookupParameter(paraname).StorageType == StorageType.String:
		value = element_type.LookupParameter(paraname).AsString()
	elif element_type.LookupParameter(paraname).StorageType == StorageType.Integer:
		value = element_type.LookupParameter(paraname).AsInteger()
	elif element_type.LookupParameter(paraname).StorageType == None:
		value = "Da Compilare"
	return value

pvp = ParameterValueProvider(ElementId(634609))
fng = FilterStringEquals()
ruleValue = 'D_ESPORTAZIONI_IFC'
fRule = FilterStringRule(pvp,fng,ruleValue,True)

filter = ElementParameterFilter(fRule)

exp_views_coll = FilteredElementCollector(doc).OfClass(View3D).WherePasses(filter).WhereElementIsNotElementType().ToElements()

exp_views_ele = []

for ev in exp_views_coll:
	if ev.IsTemplate == False:
		exp_views_ele.append(ev)

exp_views = []

for v in exp_views_ele:
	exp_views.append(v.Name)

disc_spec = ["AMB","GET","FED","TRA","IMP","STR"]

exp_view_check = []
exp_view_false = []

for v in exp_views:
    presente = False
    for i in disc_spec:
        if i in v:
            presente = True
            break
    if presente:
        exp_view_check.append("{} --> V".format(v))
    else:
        exp_view_false.append(v)

#-------------------------------------FAMIGLIE E VISTE

for ev in exp_views_ele:
	if "FED" in ev.Name:
		ex_view_fg = ev

t = Transaction(doc,'RENAME Views')
t.Start()

for v in exp_views_ele:
	if "CV" in v.Name:
		v.Name = v.Name.replace("CV","PV")
	elif "VI" in v.Name:
		v.Name = v.Name.replace("VI","PV")
		
t.Commit()


doc_el = FilteredElementCollector(doc,ex_view_fg.Id).WhereElementIsNotElementType().ToElements()

cat = []
cnamelist = []
catname = []

clean_el = []

for el in doc_el:
	try:
		if el.Category.CategoryType == CategoryType.Model and "dwg"  not in el.Category.Name and el.Category.SubCategories.Size > 0 or el.Category.CanAddSubcategory:
			clean_el.append(el)
	except:
		pass

if len(clean_el) == 0:
	forms.alert('WARNING 07_VISTA EXP FEDERATO\n\nLa vista di esportazione FED non contiene nessun elemento\n\n Correggere  le impostazioni della vista', exitscript=True)


fam_el = FilteredElementCollector(doc).OfClass(Family).ToElements()
familyname = []


t = Transaction(doc,'RENAME')
t.Start()

for fam in fam_el:
	if "CV" in fam.Name:
		fam.Name = fam.Name.replace("CV","PV")
		familyname.append(fam_el)
		
t.Commit()


#-------------------------------------OPERA

t1 = Transaction(doc,'SETPARA1')
t1.Start()

TYPE_I = []

for el in clean_el:
	TYPE_I.append(doc.GetElement(el.GetTypeId()))

UniqueType = list(set(TYPE_I))


for tp in UniqueType:
	try:
		if "CV" in Element.Name.__get__(tp):
			tp.Name = Element.Name.__get__(tp).replace("CV","PV")
	except Exception as e:
		print (e)
		pass 

	opera_el = tp.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsValueString()

	try:
		if "CV" in opera_el:
			tp.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).Set("PV")
	except:
		pass
		
t1.Commit()




#-------------------------------------CODICE WBS

t = Transaction(doc,'SETPARA2')
t.Start()

for el in clean_el:
	try:
		if "CV" in ParaInst(el,"IDE_Codice WBS"):
			Para(el,"IDE_Codice WBS").Set(ParaInst(el,"IDE_Codice WBS").replace("CV","PV"))
	except:
		pass

t.Commit()

#-------------------------------------GRUPPO ANAGRAFICA

t = Transaction(doc,'SETPARA3')
t.Start()

for el in clean_el:
	try:
		if "CV" in ParaInst(el,"IDE_Gruppo anagrafica"):
			Para(el,"IDE_Gruppo anagrafica").Set(ParaInst(el,"IDE_Gruppo anagrafica").replace("CV","PV"))
	except:
		pass

t.Commit()

print("Rename Effettuato")

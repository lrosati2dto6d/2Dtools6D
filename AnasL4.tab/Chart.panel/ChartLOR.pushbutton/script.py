"""Distribuzione Elementi secondo la loro Opera-Parte d'Opera-Elemento"""

__title__= 'Chart LOR\nAnas L-4'
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

# LISTS
# COLORS for chart.js graphs - chartCategories.randomize_colors() sometimes
# creates COLORS which are not distunguishable or visible


COLORSLOR = 10 * [
    "#FFC125",
    "#FF7F00",
    "#B0171F",
    "#00C957",
]


height = 'height:100px'
fontsize_w = 28


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

#-------------------------------------FINESTRA OUTPUT

output = script.get_output()

output.resize(900,600)

#-------------------------------------EXP.VIEWS

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


#WARNING_06-----------------NOME E PRESENZA VISTE ESPORTAZIONE

if len(exp_view_check) == len(exp_views):
	exp_view_result = "WARNING 06_VISTE ESPORTAZIONE IFC --> :white_heavy_check_mark:"
else:
	forms.alert('WARNING 06_VISTE ESPORTAZIONE IFC\n\nLe seguenti viste di esportazione non sono correttemente impostate o rinominate\n\n {}'.format(exp_view_false), exitscript=True)

#-------------------------------------ELEMENTI

for ev in exp_views_ele:
	if "FED" in ev.Name:
		ex_view_fg = ev

#WARNING_07-----------------VISTA ESPORTAZIONE FEDERATO

if ex_view_fg.get_Parameter(BuiltInParameter.VIEW_PHASE).AsValueString()=="Nuova costruzione" and ex_view_fg.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER).AsValueString()=="Mostra completo":
	exp_viewed_result = "WARNING 07_VISTA EXP FEDERATO --> :white_heavy_check_mark:"
else:
	forms.alert('WARNING 07_VISTA EXP FEDERATO\n\nLa vista di esportazione FED non ha le fasi impostate correttamente\n\n Correggere la vista con le seguenti impostazioni\n\nFiltro delle fasi = Mostra completo\nFase = Nuova costruzione', exitscript=True)

doc_el = FilteredElementCollector(doc,ex_view_fg.Id).WhereElementIsNotElementType().ToElements()

cat = []
cnamelist = []
catname = []

clean_el = []

for el in doc_el:
	try:
		if el.Category.CategoryType == CategoryType.Model and "dwg"  not in el.Category.Name and el.Category.SubCategories.Size > 0 and el.Category.CanAddSubcategory:
			clean_el.append(el)
	except:
		pass

for el in clean_el:
	type_el = doc.GetElement(el.GetTypeId())
	category_el = el.Category.Name
	opera_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsValueString()
	parteopera_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString()
	elemento_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsValueString()


lors = []

lor_counts = []

#-------------------------------------ELEMENTI PER LOR + Chart

# Conta gli elementi per categoria
for el in clean_el:
	lor_name = ParaInst(el,"IDE_LOR")

	if lor_name not in lors:
		lors.append(lor_name)
		lor_counts.append(1)

	else:
		index = lors.index(lor_name)
		lor_counts[index] += 1

labels_lors = []

for l,lc in zip(lors,lor_counts):
	somma = sum(lor_counts)
	percentage = '{} %'.format(lc * 100 / somma)
	labels_lors.append('{} : {}'.format(l,percentage))


# Crea una chart di tipo donut
chart_lor = output.make_pie_chart()

chart_lor.set_style(height)

chart_lor.options.title = {'display': True, 'text':'Distribuzione elementi per LOR', 'fontSize': fontsize_w, 'fontColor': '#000', 'fontStyle': 'bold'}

chart_lor.data.labels = labels_lors

conteggi_lor = chart_lor.data.new_dataset('numero elementi')

conteggi_lor.data = lor_counts

conteggi_lor.backgroundColor = COLORSLOR

chart_lor.draw()


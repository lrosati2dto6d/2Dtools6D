"""Distribuzione Elementi secondo la loro Opera-Parte d'Opera-Elemento"""

__title__= 'Chart Modello Dati\nAnas L-4'
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

COLORSOPERA = 10 * [
    "#1C86EE",
    "#00CED1",
    "#00CD00",
    "#FF9912",
    "#FF4040",
]


COLORS = 10 * [
    "#EE00EE",
    "#0000FF",
    "#FF8000",
    "#FF0000",
    "#228B22",
    "#00BFFF",
    "#00E5EE",
    "#00C78C",
    "#00EE76",
    "#3D9140",
    "#FFD700",
    "#FFC125",
    "#CD8500",
    "#FF4500",
    "#FF7F24",
    "#FF4500",
    "#FF6347",
    "#C67171",
    "#7D9EC0",
    "#388E8E",
    "#8E388E",
    "#71C671",
    "#FF0000",
    "#8A360F",
    "#FF8C69",
    "#228B22",
    "#e97800",
    "#a6c844",
    "#fff0e6",
    "#ffc299",
    "#ff751a",
    "#cc5200",
    "#ff6666",
    "#ffd480",
    "#b33c00",
    "#ff884d",
    "#d9d9d9",
    "#9988bb",
    "#4d4d4d",
    "#fff0e6",
    "#e97800",
    "#a6c844",
    "#ffc299",
    "#ff751a",
    "#cc5200",
    "#ff6666",
    "#ffd480",
    "#b33c00",
    "#ff884d",
    "#d9d9d9",
    "#9988bb",
    "#4d4d4d",
    "#9988bb",
    "#4d4d4d",
    "#e97800",
    "#a6c844",
    "#4d4d4d",
    "#fff0d9",
    "#ffc299",
    "#ff751a",
    "#cc5200",
    "#ff6666",
    "#ffd480",
    "#b33c00",
    "#ff884d",
    "#d9d9d9",
    "#9988bb",
    "#4d4d4d",
    "#e97800",
    "#a6c844",
    "#4d4d4d",
    "#fff0d9",
    "#ffc299",
    "#ff751a",
    "#cc5200",
    "#ff6666",
    "#ffd480",
    "#b33c00",
    "#ff884d",
    "#d9d9d9",
    "#9988bb",
    "#4d4d4d",
    "#e97800",
    "#a6c844",
    "#4d4d4d",
    "#fff0d9",
    "#ffc299",
    "#ff751a",
    "#cc5200",
    "#ff6666",
    "#ffd480",
    "#b33c00",
    "#ff884d",
    "#d9d9d9",
    "#9988bb",
    "#4d4d4d",
    "#e97800",
    "#a6c844",
    "#4d4d4d",
    "#fff0d9",
    "#ffc299",
    "#ff751a",
    "#cc5200",
    "#ff6666",
    "#ffd480",
    "#b33c00",
    "#ff884d",
    "#d9d9d9",
    "#9988bb",
    "#4d4d4d",
    "#e97800",
    "#a6c844",
    "#4d4d4d",
    "#fff0d9",
    "#ffc299",
    "#ff751a",
    "#cc5200",
    "#ff6666",
    "#ffd480",
    "#b33c00",
    "#ff884d",
    "#d9d9d9",
    "#9988bb",
    "#4d4d4d",
    "#e97800",
    "#a6c844",
]

height = 'height:80px'
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

output.resize(1200,800)

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


#-------------------------------------ELEMENTI PER OPERA + Chart

operas = []
poperas = []
eloperas = []
lors = []


opera_counts = []
popera_counts = []
elopera_counts = []
lor_counts = []


# Conta gli elementi per categoria
for el in clean_el:
	type_el = doc.GetElement(el.GetTypeId())
	opera_name = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsValueString()

	if opera_name not in operas:
		operas.append(opera_name)
		opera_counts.append(1)

	else:
		index = operas.index(opera_name)
		opera_counts[index] += 1

labels_opera = []

for o,oc in zip(operas,opera_counts):
	somma = sum(opera_counts)
	percentage = '{} %'.format(round(oc * 100 / somma),4)
	labels_opera.append('{} : {}'.format(o,percentage))



# Crea una chart di tipo donut
chart_opera = output.make_bar_chart()

chart_opera.set_style(height)

chart_opera.options.title = {'display': True, 'text':'Distribuzione elementi per Opera', 'fontSize': fontsize_w, 'fontColor': '#000', 'fontStyle': 'bold'}
chart_opera.options.legend = {'display': False}
chart_opera.data.labels = labels_opera

conteggi_opera = chart_opera.data.new_dataset('numero elementi')

conteggi_opera.data = opera_counts

conteggi_opera.backgroundColor = COLORSOPERA

chart_opera.draw()

output.print_md(	'-----------------------')

#-------------------------------------ELEMENTI PER PARTE D'OPERA + Chart


# Conta gli elementi per categoria
for el in clean_el:
	type_el = doc.GetElement(el.GetTypeId())
	popera_name = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString()

	if popera_name not in poperas:
		poperas.append(popera_name)
		popera_counts.append(1)

	else:
		index = poperas.index(popera_name)
		popera_counts[index] += 1



# Crea una chart di tipo donut
chart_popera = output.make_bar_chart()

chart_popera.set_style(height)
chart_popera.options.title = { 'align': 'start', 'display': True, 'text':'Distribuzione elementi per Parte Opera', 'fontSize': fontsize_w, 'fontColor': '#000', 'fontStyle': 'bold'}
chart_popera.options.legend = {'display': False}
chart_popera.data.labels = poperas

conteggi_popera = chart_popera.data.new_dataset('numero elementi')

conteggi_popera.data = popera_counts

conteggi_popera.backgroundColor = COLORS

chart_popera.draw()

output.print_md(	'-----------------------')

#-------------------------------------ELEMENTI PER CODICE ELEMENTO + Chart

# Conta gli elementi per categoria
for el in clean_el:
	type_el = doc.GetElement(el.GetTypeId())
	elopera_name = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsValueString()

	if elopera_name not in eloperas:
		eloperas.append(elopera_name)
		elopera_counts.append(1)

	else:
		index = eloperas.index(elopera_name)
		elopera_counts[index] += 1

# Crea una chart di tipo donut
chart_elopera = output.make_bar_chart()

chart_elopera.set_style(height)
chart_elopera.options.title = {'display': True, 'text':'Distribuzione elementi per Codice Elemento', 'fontSize': fontsize_w, 'fontColor': '#000', 'fontStyle': 'bold'}
chart_elopera.options.legend = {'display': False}

chart_elopera.data.labels = eloperas

conteggi_elopera = chart_elopera.data.new_dataset("CODICE ELEMENTO")

conteggi_elopera.data = elopera_counts

conteggi_elopera.backgroundColor = COLORS

chart_elopera.draw()


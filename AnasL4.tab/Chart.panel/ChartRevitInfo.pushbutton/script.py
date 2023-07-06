"""Distribuzione Elementi secondo la loro Categoria-Workset-Fase"""

__title__= 'Chart Revit Info\nAnas L-4'
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

COLORSWORK = 10 * [
    "#EE00EE",
    "#0000FF",
    "#FF8000",
    "#FF0000",
    "#228B22",
]

COLORSFASE = 10 * [
    "#FFD700",
    "#FF0000",
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
]

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
		if el.Category.CategoryType == CategoryType.Model and "dwg"  not in el.Category.Name and el.Category.SubCategories.Size > 0 or el.Category.CanAddSubcategory:
			clean_el.append(el)
	except:
		pass

for el in clean_el:
	catname.append(el.Category.Name)
	cat.append(el.Category)
	cnamelist.append(el)


#-------------------------------------ELEMENTI PER CATEGORIA + Chart

categories = []
cat_counts = []

# Conta gli elementi per categoria
for el in clean_el:
	category = el.Category
	category_name = category.Name
	if category_name not in categories:
		categories.append(category_name)
		cat_counts.append(1)
	else:
		index = categories.index(category_name)
		cat_counts[index] += 1

# Crea una chart di tipo donut
chart_cate = output.make_pie_chart()

chart_cate.set_style('height:150px')
chart_cate.options.title = {'display': True, 'text':'Distribuzione elementi per Categoria', 'fontSize': 24, 'fontColor': '#000', 'fontStyle': 'bold'}

chart_cate.data.labels = categories

conteggi_cat = chart_cate.data.new_dataset('numero elementi')

conteggi_cat.data = cat_counts

conteggi_cat.backgroundColor = COLORS

chart_cate.draw()

output.print_md(	'-----------------------')

#-------------------------------------ELEMENTI PER WORKSET + Chart

worksets = []
wor_counts = []

for el in clean_el:
	workset_name = el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString()
	if workset_name not in worksets:
		worksets.append(workset_name)
		wor_counts.append(1)
	else:
		index = worksets.index(workset_name)
		wor_counts[index] += 1

# Crea una chart di tipo donut
chart_work = output.make_pie_chart()

chart_work.set_style('height:150px')
chart_work.options.title = {'display': True, 'text':'Distribuzione elementi per Workset', 'fontSize': 24, 'fontColor': '#000', 'fontStyle': 'bold'}

chart_work.data.labels = worksets

conteggi_work = chart_work.data.new_dataset('numero elementi')

conteggi_work.data = wor_counts

conteggi_work.backgroundColor = COLORSWORK

chart_work.draw()

output.print_md(	'-----------------------')

#-------------------------------------ELEMENTI PER FASE + Chart

phases = []
pha_counts = []

for el in clean_el:
	pha_name = el.get_Parameter(BuiltInParameter.PHASE_CREATED).AsValueString()
	if pha_name not in phases:
		phases.append(pha_name)
		pha_counts.append(1)
	else:
		index = phases.index(pha_name)
		pha_counts[index] += 1

# Crea una chart di tipo donut
chart_pha = output.make_pie_chart()

chart_pha.set_style('height:150px')
chart_pha.options.title = {'display': True, 'text':'Distribuzione elementi per Fase di Creazione', 'fontSize': 24, 'fontColor': '#000', 'fontStyle': 'bold'}

chart_pha.data.labels = phases

conteggi_pha = chart_pha.data.new_dataset('numero elementi')

conteggi_pha.data = pha_counts

conteggi_pha.backgroundColor = COLORS

chart_pha.draw()

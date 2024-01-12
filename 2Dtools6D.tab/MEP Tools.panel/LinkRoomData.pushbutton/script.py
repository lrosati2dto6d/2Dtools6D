"""Utilizzo: 1-Selezionare il Link in cui risiedono le Room  2-Selezionare i parametri delle Room da trasferire  3- Scriviere un Parametro di Testo da popolare con il valore selezionato"""

__title__= 'Link Room\nData'
__author__= 'Luca Rosati'

import clr

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import *

clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(Revit.Elements)

clr.AddReference('RevitServices')
from System.Collections.Generic import *

from pyrevit import forms
from pyrevit import script

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

#-------------------------------------FINESTRA OUTPUT

output = script.get_output()

output.resize(1200,800)

work_ele=[]

#----------------------FILTER POINT INSTANCE

doc_el_tot = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

for el in doc_el_tot:
	try:
		if el.Category.CategoryType == CategoryType.Model and "dwg" not in el.Category.Name and el.Category.SubCategories.Size > 0 and el.Category.CanAddSubcategory and "Detail" not in el.Category.Name and "Line" not in el.Category.Name:
			work_ele.append(el)
	except:
		pass

work_ele_point = []
work_point = []
categoryN = set()

for wl in work_ele:
	try:
		if wl.Location.GetType()== LocationPoint:
			work_ele_point.append(wl)
			work_point.append(wl.Location.Point)
	except:
		pass

#----------------------SELECT LINK

lnks = FilteredElementCollector(doc).OfClass(RevitLinkInstance)

doclnk=[]
doclnkn=[]

for i in lnks:
	doclnk.append(i.GetLinkDocument())
	try:
		doclnkn.append(i.GetLinkDocument().Title)
	except:
		pass

if len(doclnk) == 0:
	forms.alert('There are no links in the current document, please upload at least one', exitscript=True)

str_linkname = forms.ask_for_one_item(
    doclnkn,
    default = doclnkn[0],
    prompt='Select the Link where are the Rooms From List',
    title='ROOM LINK SELECTOR')

ldoc = None
linkinst = []
for i,j,ins in zip(doclnk,doclnkn,lnks):
	if str_linkname == j:
			ldoc=i
			linkinst=ins

if ldoc == None:
	script.exit()

roomsinst = FilteredElementCollector(ldoc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements()

if len(roomsinst)==0:
	forms.alert('Non sono presenti Rooms nel documento selezionato', exitscript=True)

p_par = roomsinst[0].GetOrderedParameters()
p_par_n = []
for p in p_par:
	p_par_n.append(p.Definition.Name)

sortlist = sorted(p_par_n)

res = forms.SelectFromList.show(sortlist,
                                multiselect = True,
                                title ='Select Parameters to Concatenate Values',
                                group_selector_title ='Link Room Data')

if len(res) == 0:
	forms.alert('Scegliere almeno un Parametro', exitscript=True)
else:
    forms.toast(
        '-'.join(res),
        title="Result of Selected Parameters",
    )

list_in = range(0,len(res),1)
list_in_str = []

for listr in list_in:
	list_in_str.append('{}'.format(listr))

appear_list = []

for r,n in zip(res,list_in):
	appear_list.append('{}={}'.format(r,n))

para_ord = forms.ask_for_string(
    default=','.join(list_in_str),
    prompt="Se vuoi cambiare il valore risultante, cambia l'ordine degli indici in\nmodo da ricomporre la concatenazione\nricordati di inserire una virgola tra un indice e un altro",
    title='-'.join(appear_list)
)

x = para_ord.split(",")
a = []
for i in x:
	a.append(int(i))

list_res = []

for r,y in zip(res,a):
	list_res.append(res[y])

para_text = forms.ask_for_string(
    default='Parameter',
    prompt='Enter Parameter name:',
    title='Link Room Data'
)

if para_text == 'Parameter' or para_text == None:
	forms.alert('Scrivere un Parametro Testo\nin cui trasferire il valore', exitscript=True)

for e in work_ele_point:
	try:
		e.LookupParameter(para_text)
	except:
		forms.alert('Il Parametro - {} - non esiste in tutti gli elementi coinvolti\n Scegliere un altro parametro'.format(para_text), exitscript=True)
	
#----------------------CHECK POINT IN ROOM

roomsl = []
pointsl=[]
elem =[]

for roomi in roomsinst:
	pointslist =[]
	elepoint = []
	roomsl.append(roomi)
	for pointi,ele in zip(work_point,work_ele_point):
		if roomi.IsPointInRoom(pointi):
			pointslist.append(pointi)
			elepoint.append(ele)
	pointsl.append(pointslist)
	elem.append(elepoint)

t_Compilare = Transaction(doc,"Inserimento Roominfo")
t_Compilare.Start()

final_value = []
for r in roomsl:
	f_list = []
	for re in list_res:
		if len(list_res) != 0:
			f_list.append(r.LookupParameter(re).AsValueString())
	final_value.append('-'.join(f_list))

for r,e,f in zip(roomsl,elem,final_value):
	for el in e:
		el.LookupParameter(para_text).Set(f)

t_Compilare.Commit()
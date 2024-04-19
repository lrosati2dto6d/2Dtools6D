"""Select Elements based on a Range of Parameter Values"""

__title__= 'Select by\n Value Range'
__author__= 'Luca Rosati'

import System
import clr
import sys
sys.path.append('C:\Program Files (x86)\IronPython 2.7\Lib')

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import *

clr.AddReference('System')
from System.Collections.Generic import *

clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(Revit.Elements)

clr.AddReference('RevitServices')

from pyrevit import DB
from pyrevit.framework import List
from pyrevit import forms
from pyrevit import script

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument

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

output.resize(600,300)

categories =doc.Settings.Categories

model_cat = []
model_cate = []
for c in categories:
	if c.CategoryType == CategoryType.Model:
		if "dwg"  not in c.Name and c.SubCategories.Size > 0 or c.CanAddSubcategory:
			model_cat.append(c.Name)
			model_cate.append(c)

sortlist = sorted(model_cat)

value = forms.SelectFromList.show(
        {'All Categories': sortlist},
        title='Select Categories',
        multiselect=False
    )

category =None
namer = []
categoriesId = []
bic = []

for ci in model_cate:
	namer.append(ci.Name)

for n,c in zip(namer,model_cate):
	if n == value:
            category = c

if category == None:
	script.exit()

categoriesId = category.Id
bic= (System.Enum.ToObject(BuiltInCategory, category.Id.IntegerValue))

#print(category.Name,categoriesId,bic)

catinst = FilteredElementCollector(doc).OfCategory(bic).WhereElementIsNotElementType().ToElements()

if len(catinst)==0:
	forms.alert('Non sono presenti elementi della categoria selezionata', exitscript=True)

if len(catinst)==1:
	forms.alert('Scegliere una categoria con almeno due elementi', exitscript=True)

p_par =  catinst[0].GetOrderedParameters()
p_par_n = []

for p in p_par:
	if p.StorageType == StorageType.Double:
		p_par_n.append(p.Definition.Name)

sortlistp = sorted(p_par_n)

res = forms.SelectFromList.show(sortlistp,
                                multiselect = False,
                                title ='Select Parameter To Evaluate Range',
                                group_selector_title ='Isolate Value Range')



llo = catinst[0].LookupParameter(res).GetUnitTypeId()

bool = catinst[0].LookupParameter(res).Definition.GetDataType().TypeId


prop_infos = clr.GetClrType(DB.SpecTypeId).GetProperties()
dict_forgeTypeId_id = [p.GetGetMethod().Invoke(None, None).TypeId for p in prop_infos]
dict_forgeTypeId = [p.GetGetMethod().Invoke(None, None) for p in prop_infos]
dict_forgeTypeId_name = [p.Name for p in prop_infos]

unittypec = None

for pti,ptiid in zip(dict_forgeTypeId,dict_forgeTypeId_id):
	if bool == ptiid:
		unittypec = pti


def ConvertAllUnits(num,forgetyppara):
	unitType = forgetyppara
	currentDisplayUnits = doc.GetUnits().GetFormatOptions(unitType).GetUnitTypeId()
	return UnitUtils.ConvertFromInternalUnits(float(num),currentDisplayUnits)

valuesconv = []

llu = UnitUtils.GetTypeCatalogStringForUnit(llo)
#lli = UnitUtils.GetValidDisplayUnits(doc.GetUnits().GetFormatOptions(unittypec).GetUnitTypeId())

list_ele = []

for inst in catinst:
	try:
		valuesconv.append(ConvertAllUnits(ParaInst(inst,res),unittypec))
		list_ele.append(inst)
	except:
		pass

min_v = min(valuesconv)
max_v = max(valuesconv)

res_min = forms.ask_for_number_slider(default=None, min=min_v, max=max_v, interval=0.1, prompt="Select the Minimun value of Range", title="Unit: {}".format(llu))

res_max = forms.ask_for_number_slider(default=None, min=res_min, max=max_v, interval=0.1, prompt="Select the Maximun value of Range, Min Value = {}".format(res_min), title="Unit: {}".format(llu))

inst_result=[]
inst_result_id=[]

inst_result_inv=[]
inst_result_inv_id=[]

for v,i in zip(valuesconv,list_ele):
	if res_min <= v <= res_max: 
		inst_result.append(i)
		inst_result_id.append(i.Id)
	else:
		inst_result_inv.append(i)
		inst_result_inv_id.append(i.Id)

if len(inst_result)==0:
	forms.alert('Non sono presenti Elementi aventi il valore nel Range considerato', exitscript=True)


selected_option1 = forms.CommandSwitchWindow.show(
    ['No', 'Yes'],
     message='Inverse Selection?',
)


if selected_option1 == "No":
	collection = List[ElementId](inst_result_id)
	select = uidoc.Selection.SetElementIds(collection)

	output = script.get_output()

	output.print_md(	'# CATEGORY : {}'.format(value))

	output.print_md(	'# PARAMETER : {}'.format(res))

	output.print_md(	'# VALUE RANGE {} : [ {} - {} ]'.format(llu,res_min,res_max))

	output.print_md(	'##\tELEMENTS SELECTED : n.{} of n.{} Total'.format(len(inst_result_id),len(list_ele)))

	output.print_md(	'##\tInverse Selection? {}'.format(selected_option1))

else:
	collection = List[ElementId](inst_result_inv_id)
	select = uidoc.Selection.SetElementIds(collection)

	output = script.get_output()

	output.print_md(	'# CATEGORY : {}'.format(value))

	output.print_md(	'# PARAMETER : {}'.format(res))

	output.print_md(	'# VALUE RANGE {} : [ {} - {} ]'.format(llu,res_min,res_max))

	output.print_md(	'##\tELEMENTS SELECTED : n.{} of n.{} Total'.format(len(inst_result_inv_id),len(list_ele)))

	output.print_md(	'##\tInverse Selection? {}'.format(selected_option1))
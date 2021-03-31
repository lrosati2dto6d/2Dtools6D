"""Select Elements based on the value of the indicated parameter and its Category"""

__title__= 'Select By\nParamater Value'
__author__= 'Luca Rosati'

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

from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document
uidoc =  __revit__.ActiveUIDocument


with forms.WarningBar(title='Pick Element:'):
	sel = uidoc.Selection
	selected = sel.PickObject(ObjectType.Element)

to_element = doc.GetElement(selected)

p_par = to_element.GetOrderedParameters()
p_par_n = []
for p in p_par:
	p_par_n.append(p.Definition.Name)

sortlist = sorted(p_par_n)

value = forms.ask_for_one_item(
    sortlist,
    default= sortlist[0],
    prompt='Select Parameter',
    title='Select By Paramater ')

selected_option = forms.CommandSwitchWindow.show(
    ['Yes', 'No'],
     message='Select in Active View?',
)
selected_option1 = forms.CommandSwitchWindow.show(
    ['Yes', 'No'],
     message='Inverse Selection?',
)

el_cat = to_element.Category.Id

selected_bc = None

builtin_cats = System.Enum.GetValues(BuiltInCategory)

bc_lst = []
bcid_lst = []

for bc in builtin_cats:
	try:
		bc_lst.append(bc)
		bcid_lst.append(ElementId(bc))
	except:
		pass

for i, y in zip(bc_lst, bcid_lst):
	if el_cat == y:
		selected_bc = i

activeview_toggle=None

if selected_option=='Yes':
	activeview_toggle=True
else:
	activeview_toggle=False


if activeview_toggle:
	collectordoc = FilteredElementCollector(doc,doc.ActiveView.Id)
else:
	collectordoc = FilteredElementCollector(doc)

fam_insts = collectordoc.OfCategory(selected_bc).WhereElementIsNotElementType().ToElements()

selected_ele_param = to_element.LookupParameter(value)

value_par = None 
values_other = []

try:
	if selected_ele_param.StorageType == StorageType.Double:
		value_par = Parameter.AsDouble(selected_ele_param)
	elif selected_ele_param.StorageType == StorageType.ElementId:
		value_par = Parameter.AsElementId(selected_ele_param)
	elif selected_ele_param.StorageType == StorageType.String:
		value_par = selected_ele_param.AsString()
	elif selected_ele_param.StorageType == StorageType.Integer:
		value_par = Parameter.AsInteger(selected_ele_param)
	elif selected_ele_param.StorageType == None:
		selected_ele_param.StorageType = None
except:
	pass

try:
	for i in fam_insts:
		if i.LookupParameter(value).StorageType == StorageType.Double:
			values_other.append(i.LookupParameter(value).AsDouble())
		elif i.LookupParameter(value).StorageType == StorageType.ElementId:
			values_other.append(i.LookupParameter(value).AsElementId())
		elif i.LookupParameter(value).StorageType == StorageType.String:
			values_other.append(i.LookupParameter(value).AsString())
		elif i.LookupParameter(value).StorageType == StorageType.Integer:
			values_other.append(i.LookupParameter(value).AsInteger())
		elif i.LookupParameter(value).StorageType == None:
			values_other.append(None)
except:
	pass

matched_fam_id = []
matched_fam = []

if selected_option1=='Yes':
	inverse_toggle=False
else:
	inverse_toggle=True

if inverse_toggle:
	for i, j in zip(values_other, fam_insts):
		if i == value_par:
			matched_fam_id.append(j.Id)
			matched_fam.append(j)
else:
	for i, j in zip(values_other, fam_insts):
		if i != value_par:
			matched_fam_id.append(j.Id)
			matched_fam.append(j)

collection = List[ElementId](matched_fam_id)
select = uidoc.Selection.SetElementIds(collection)

from pyrevit import script

output = script.get_output()

output.print_md(	'# CATEGORY: {}'.format(to_element.Category.Name))

output.print_md(	'# PARAMETER MATCHED : {}'.format(value))

output.print_md(	'# VALUE MATCHED : "{}"'.format(value_par))

output.print_md(	'##\tELEMENTS SELECTED: n.{}'.format(len(matched_fam)))

output.print_md(	'##\tSelected in Active View? {} - Inverse Selection? {}'.format(selected_option,selected_option1))

if matched_fam== []:
	print("The Parameter not Exists in Selected Element")
else:
	for idx, elid in enumerate(matched_fam):
		print('{}: {} - {}'.format(idx+1,elid.Name,output.linkify(elid.Id)))


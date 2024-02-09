"""Creates a table where all woksets associated with links in the project are displayed"""

__title__= 'Link-Workset\nInfo'
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
import RevitServices
from RevitServices.Persistence import DocumentManager

from collections import defaultdict
from pyrevit import revit, DB
from pyrevit import HOST_APP
from pyrevit.framework import List
from pyrevit import coreutils
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

#-------------------------------------FINESTRA OUTPUT

output = script.get_output()

output.resize(800,800)

lnks = FilteredElementCollector(doc).OfClass(RevitLinkInstance)

if forms.check_workshared() == False:
   script.exit()

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

work_par = []

for l in lnks:
	work_par.append(l.LookupParameter("Workset").AsValueString())


output = script.get_output()

output.print_md(	'#Current List of Link-Workset')

for n,w in zip (doclnkn,work_par):
	output.print_md(	'#[{}] ----> [{}]'.format(n,w))

"""
else:
	collection = List[ElementId](inst_result_inv_id)
	select = uidoc.Selection.SetElementIds(collection)

	output = script.get_output()

	output.print_md(	'# CATEGORY : {}'.format(value))

	output.print_md(	'# PARAMETER : {}'.format(res))

	output.print_md(	'# VALUE RANGE {} : [ {} - {} ]'.format(llu,res_min,res_max))

	output.print_md(	'##\tELEMENTS SELECTED : n.{} of n.{} Total'.format(len(inst_result_inv_id),len(list_ele)))

	output.print_md(	'##\tInverse Selection? {}'.format(selected_option1))

"""
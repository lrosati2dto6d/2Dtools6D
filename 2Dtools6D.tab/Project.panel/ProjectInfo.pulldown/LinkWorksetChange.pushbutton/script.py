"""Select Workset and apply to a list of Links"""

__title__= 'Link-Workset\nChange'
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


workset_list = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset)

wor_name=[]

for wor in workset_list:
	wor_name.append(wor.Name)


value = forms.SelectFromList.show(
        {'All Links Loaded': doclnkn},
        title='Select Links',
        multiselect=True
    )

if value == None:
	forms.alert('There are no links selected, please select at least one', exitscript=True)

ldoc =[] 
linkinst = []
for i,j,ins in zip(doclnk,doclnkn,lnks):
	if j in value:
		ldoc.append(i)
		linkinst.append(ins)


workset_n = forms.ask_for_one_item(
    wor_name,
    default = wor_name[0],
    prompt='Select Workset to Assign',
    title='Workset Selector')

work_set = None

for worl in workset_list:
	if worl.Name == workset_n:
		work_set = worl

if work_set == None:
	forms.alert('There are no Workset selected, please select at least one', exitscript=True)

t_work = Transaction(doc,"Set workset")
t_work.Start()

for l in linkinst:
	l.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM ).Set(work_set.Id.IntegerValue)
	doc.GetElement(l.GetTypeId()).get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM ).Set(work_set.Id.IntegerValue)

t_work.Commit()


work_par_i = []

for l in lnks:
	work_par_i.append(l.LookupParameter("Workset").AsValueString())

output = script.get_output()

output.print_md(	'#New List of Link-Workset')

for n,w in zip (doclnkn,work_par_i):
	output.print_md(	'#[{}] ----> [{}]'.format(n,w))
"""Transfers any string parameter value from the pipe element to its corresponding insulator host element"""

__title__= "Transfer Parameter\nto Host Insulation"
__author__= "Luca Rosati"

from collections import defaultdict
from pyrevit import HOST_APP
from pyrevit.framework import List
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script

import clr

import System
from System import Array
from System.Collections.Generic import *

clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

import Autodesk 
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

elems_pipes = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
elems_ducts = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()

if len(elems_pipes)!= 0:
	param_pipes = elems_pipes[0].GetOrderedParameters()
para_name_list_pipes = []

if len(elems_ducts)!= 0:
	param_ducts = elems_ducts[0].GetOrderedParameters()
para_name_list_ducts = []

type = []

for p in param_pipes:
	if p.StorageType == StorageType.String and p.IsReadOnly == False:
		para_name_list_pipes.append(p.Definition.Name)

for p in param_pipes:
	if p.StorageType == StorageType.String and p.IsReadOnly == False:
		para_name_list_ducts.append(p.Definition.Name)


res= forms.SelectFromList.show(
        {'All': para_name_list_pipes},
        title='Parameters Selector',
        group_selector_title='All:',
        multiselect=True
    )

def tolist(input):       
    result = input if isinstance(input, list) else [input]
    return result

def uwlist(input):
    result = input if isinstance(input, list) else [input]
    return UnwrapElement(input)

all_pipeinsulation = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeInsulations).WhereElementIsNotElementType().ToElements() 
all_ductinsulation = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctInsulations).WhereElementIsNotElementType().ToElements()
faminsts = all_pipeinsulation #+all_ductinsulation

host_list = [] 

for item in faminsts:
	host_list.append(doc.GetElement(item.HostElementId))

par_list = [tolist(res)]

par_values=[]
for host in host_list:
	for i in par_list:
		par_values.append([host.LookupParameter(z).AsString() for y in par_list for z in y])

host_par=[]
for item in faminsts:
	for i in par_list:
		host_par.append([item.GetParameters(x) for x in i])

def Flatten(list):
	listout= []
	for x in list:
		for i in x:
			listout.append(i)
	return listout

def Flattentot(list):
	listout=[]
	for x in list:
		for i in x:
			listout.append(i)
	return listout
	
flat_host_par= Flatten(host_par)
flat_values=Flattentot(par_values)

para_set = []

t = Transaction(doc,"Set Parameters")

t.Start()
try:
	for i,j in zip(flat_host_par,flat_values):
		if j != None:
			[y.Set(j) for y in i]
except:
	t.RollBack()
else:
	t.Commit()


output = script.get_output()
output.set_width(200)

print('Success Elements Trasfer for {} elements'.format(len(faminsts)),flat_values)



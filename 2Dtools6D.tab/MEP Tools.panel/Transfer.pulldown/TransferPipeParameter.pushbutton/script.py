"""Transfers any string parameter value from the pipe element to its corresponding insulation host element - All Categories of Piping must have same Parameters"""

__title__= "Transfer Parameter\nto Pipe Insulation"
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

elems_pipes = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()

if len(elems_pipes)!= 0:
	param_pipes = elems_pipes[0].GetOrderedParameters()
else:
	forms.alert('There must be at least one Pipe with insulation', exitscript=True)

para_name_list_pipes = []

for p in param_pipes:
	if p.StorageType == StorageType.String and p.IsReadOnly == False:
		para_name_list_pipes.append(p.Definition.Name)

res= forms.SelectFromList.show(
        {'All': para_name_list_pipes},
        title='Parameters Selector',
        group_selector_title='All:',
        multiselect=True
    )

def tolist(input):       
    result = input if isinstance(input, list) else [input]
    return result


all_pipeinsulation = FilteredElementCollector(doc, doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_PipeInsulations).WhereElementIsNotElementType().ToElements() 

faminsts = all_pipeinsulation

host_list = [] 

for item in faminsts:
	host_list.append(doc.GetElement(item.HostElementId))

par_list = res

p_val_Pipe = []
p_par_ins = []
pipe_no = []

for h,f in zip(host_list,faminsts):
	for p in par_list:
		try:
			p_val_Pipe.append(h.LookupParameter(p).AsValueString())
			p_par_ins.append(f.LookupParameter(p))
		except:
			try:
				if h.Id not in pipe_no:
					pipe_no.append(h.Id)
			except:
				pass

t = Transaction(doc,"Set Parameters")

t.Start()

for i,j in zip(p_par_ins,p_val_Pipe):
	try:
		i.Set(j)
	except:
		pass

t.Commit()

output = script.get_output()
output.set_height(600)

output.print_md('#\tSuccess Transfer for Insulation Pipe Elements')
output.print_md("##\tTransfer Completed for Parameters: {}" .format(res))


if len(pipe_no) != 0:
	output.print_md("##\tTransfer Not Completed for These Elements:")

	for i in pipe_no:
		print(i)

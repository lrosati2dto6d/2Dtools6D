"""Get Info about Project Parameters and Category assigned"""

__title__= "Get Project\nParameters Info"
__author__= "Luca Rosati"

import System
import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
import Autodesk

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager


from collections import defaultdict
from pyrevit import HOST_APP
from pyrevit.framework import List
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

names = []
groups = []
pgroup = []
ptype = []
isvis = []
elements = []
guids = []
isinst = []
bics = []
categories = []
iterator = doc.ParameterBindings.ForwardIterator()

while iterator.MoveNext():
	groups.append(iterator.Key.VariesAcrossGroups)
	names.append(iterator.Key.Name)
	pgroup.append(iterator.Key.ParameterGroup)
	try:
		ptype.append(iterator.Key.ParameterType)
	except:
		pass
	isvis.append(iterator.Key.Visible)
	elem = doc.GetElement(iterator.Key.Id)
	elements.append(elem)
	if elem.GetType().ToString() == 'Autodesk.Revit.DB.SharedParameterElement':
		guids.append(elem.GuidValue)
	else:
		guids.append(None)
	if iterator.Current.GetType().ToString() == 'Autodesk.Revit.DB.InstanceBinding':
		isinst.append("Instance")
	else:
		isinst.append("Type")
		
	thesecats = []
	builtincats = []
	for cat in iterator.Current.Categories:
		try:
			thesecats.append(cat.Name)
		except:
			thesecats.append(None)
		builtincats.append(System.Enum.ToObject(BuiltInCategory, cat.Id.IntegerValue))
	categories.append(thesecats)

numbers = range(0,len(names),1)

from pyrevit import script

output = script.get_output()

output.print_md(	'# **FILE NAME: {}**\n\tNumber of Parameters = {}'.format(doc.Title,len(names)))
if len(ptype) != 0:
	for n,pn,ii,gr,ty,gu,vr,iv,cat in zip(numbers,names,isinst,pgroup,ptype,guids,groups,isvis,categories):
		output.print_md(	'## **[{}] PARAMETER NAME: {}**'.format(n,pn))
		print('\tIS INSTANCE: {}\n\tGROUP: {}\n\tTYPE: {}\n\tGUID CODE: {}\n\tVARY BY MODEL GROUP: {}\n\tIS VISIBLE: {}\n\tAPPLIED TO: {}\n'.format(ii,gr,ty,gu,vr,iv,[x for x in cat]))
else:
	for n,pn,ii,gr,gu,vr,iv,cat in zip(numbers,names,isinst,pgroup,guids,groups,isvis,categories):
		output.print_md(	'## **[{}] PARAMETER NAME: {}**'.format(n,pn))
		print('\tIS INSTANCE: {}\n\tGROUP: {}\n\tGUID CODE: {}\n\tVARY BY MODEL GROUP: {}\n\tIS VISIBLE: {}\n\tAPPLIED TO: {}\n'.format(ii,gr,gu,vr,iv,[x for x in cat]))



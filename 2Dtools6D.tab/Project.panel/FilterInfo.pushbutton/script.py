""" Extract info regarding Filters in the current document """

__title__= "Filters\nInfo"

import System
import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
import Autodesk

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType


clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager

from System.Collections.Generic import *

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from pyrevit import forms
from pyrevit import HOST_APP
from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script

from collections import defaultdict
from pyrevit import HOST_APP
from pyrevit.framework import List
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument

view_fil = set()
fil_doc = set()
fil_name = []
fil_del = []
fil_del_name =[]
fil_view = []
view_vie = []

doc_view = FilteredElementCollector(doc).OfClass(View).ToElements()

doc_fil = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()


for f in doc_fil:
	fil_doc.add(f.Id)
	fil_name.append(f.Name)
	fil_view.append(doc.GetElement(f.OwnerViewId))

for v in doc_view:
	if v.AreGraphicsOverridesAllowed():
		fil_v = v.GetFilters()
		for f in fil_v:
			i=[]
			if f != None:
				i.append(v)
				view_fil.add(f)
			view_vie.append(i)

for f in fil_doc:
	if f not in view_fil:
		fil_del.append(f)
		fil_del_name.append(doc.GetElement(f).Name)
n_fil = "{}".format(len(doc_fil))

if len(fil_del) == 0:
	n_fil_del = "No Filter Unused"
else:
	n_fil_del = "N.Filter Unused = {} - Filters Name: {}".format(len(fil_del),fil_del_name)

#Categorie Applicate ai Filtri

fil_categories = []

for fil in doc_fil:
	list_emp = []
	cat_label = fil.GetCategories()
	for c in cat_label:
		list_emp.append(LabelUtils.GetLabelFor(System.Enum.ToObject(BuiltInCategory, c.IntegerValue)))
	fil_categories.append(list_emp)

#Parametri applicate ai Filtri

fil_rules = []

for fil in doc_fil:
	d = []
	fil_rules.append(fil.GetElementFilterParameters())

hashset = []

for fhs in fil_rules:
	y = []
	for h in fhs:
		if h.IntegerValue < 0:
			y.append(LabelUtils.GetLabelFor(System.Enum.ToObject(BuiltInParameter, h.IntegerValue)))
		else:
			y.append(doc.GetElement(h).Name)
	hashset.append(y)


numbers = range(0,len(fil_name),1)

output = script.get_output()


output.print_md(	'# FILTER INFO: {}'.format(doc.Title))
output.print_md(	'##Total Number of Filters = {}'.format(n_fil))
output.print_md(	'###{}'.format(n_fil_del))

for n,pn,pf,cf in zip(fil_name,numbers,hashset,fil_categories):
	output.print_md(	'## **[{}] FILTER NAME: {}**'.format(pn,n))
	print('\tApplied Parameters: {}\n\tApplied Categories: {}'.format(pf,cf))
"""Select Elements based on Rectangular selection and Category """

__title__= 'Rectangular\nBy Category'
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

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument

categories =doc.Settings.Categories

model_cat = []


for c in categories:
	if c.CategoryType == CategoryType.Model:
		if "dwg"  not in c.Name and c.SubCategories.Size > 0 or c.CanAddSubcategory:
			model_cat.append(c.Name)

sortlist = sorted(model_cat)

value = forms.ask_for_one_item(
    sortlist,
    default= sortlist[0],
    prompt='Select Category',
    title='Rectangular Selection')

class MySelectionFilter(ISelectionFilter):
	def __init__(self):
		pass
	def AllowElement(self, element):
		if element.Category.Name == value:
			return True
		else:
			return False
	def AllowReference(self, element):
		return False

selection_filter = MySelectionFilter()

output = uidoc.Selection.PickElementsByRectangle(selection_filter)

outputID = []

for i in output:
	outputID.append(i.Id)

collection = List[ElementId](outputID)

select = uidoc.Selection.SetElementIds(collection)
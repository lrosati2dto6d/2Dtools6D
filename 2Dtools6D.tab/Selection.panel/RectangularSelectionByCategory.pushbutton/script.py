"""Select Elements based on Rectangular selection and Category """

__title__= 'Rectangular Selection \nBy Category'
__author__= 'Luca Rosati'

import clr
import System

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import *

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *

from rpw.ui.forms import TextInput
value = TextInput('Category', default="Walls")

from pyrevit import forms


doc = __revit__.ActiveUIDocument.Document
uidoc =  __revit__.ActiveUIDocument


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



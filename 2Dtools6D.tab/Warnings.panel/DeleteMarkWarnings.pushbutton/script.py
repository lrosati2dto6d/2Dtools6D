"""Delete Mark Parameter Value to avoid Warnings of Duplicate Values"""

__title__= 'Fix All\nMark Warnings'
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

from collections import defaultdict
from pyrevit import HOST_APP
from pyrevit.framework import List
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc =  __revit__.ActiveUIDocument


categories = doc.Settings.Categories

model_cat = []
mod_c=[]

for c in categories:
	if c.CategoryType == CategoryType.Model:
		if "dwg"  not in c.Name and c.SubCategories.Size > 0 or c.CanAddSubcategory:
			model_cat.append(c.Name)
			mod_c.append(c)
			
sortlist = sorted(model_cat)


res = forms.SelectFromList.show(
        {'Mark Automatic Increment Categories': ['Doors','Curtain Panels','Mechanical Equipment','Plumbing Fixtures','Sprinklers','Windows','Electrical Equipment', 'Electrical Fixtures','Lighting Fixtures'],
		},
        title='Categories Selector',
        multiselect=True
    )

category =[]
namer = []

for ci in mod_c:
	namer.append(ci.Name)

for n,c in zip(namer,mod_c):
	for r in res:
		if n == r:
			category.append(c)

categoriesId = []
bic = []
for c in category:
	categoriesId.append(c.Id)
	bic.append(System.Enum.ToObject(BuiltInCategory, c.Id.IntegerValue))

def filcategories(document,Category): #Filtra tutti gli elementi in base a una lista di categorie
	if isinstance(Category, list):
		categoriesCollector = []
		for nId in categoriesId:
			categoriesCollector.append(FilteredElementCollector(document).OfCategoryId(nId).WhereElementIsNotElementType().ToElements())
		return categoriesCollector

outputelem = filcategories(doc,category)

parametermark=[]

for o in outputelem:
	for x in o:
		parametermark.append(x.get_Parameter(BuiltInParameter.DOOR_NUMBER))


if not forms.alert('This operation involves changes to the parameters. Click "OK" Only if you are sure it is the right thing. In anycase, Make sure your models are saved and synced. '
                   'Hit OK to continue...', cancel=True):
    script.exit()

t = Transaction(doc,"Set Parameters")

t.Start()
try:
	for p in parametermark:
		p.Set('')

except:
	t.RollBack()
else:
	t.Commit()

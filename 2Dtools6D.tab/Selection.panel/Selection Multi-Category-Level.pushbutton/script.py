"""Select Elements based on Multi-Categories and Level"""

__title__= 'Multi-Category\nOn Level'
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
from pyrevit import HOST_APP
from pyrevit.framework import List
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument

levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()

level_ns =[]
for l in levels:
	level_ns.append(l.Name)

level_n = forms.ask_for_one_item(
    level_ns,
    default = level_ns[0],
    prompt='Select Level',
    title='Level Selector')

for lev in levels:
	if lev.Name == level_n:
		level = lev

filterlevel = ElementLevelFilter(level.Id)

categories = doc.Settings.Categories

model_cat = []
mod_c=[]

for c in categories:
	if c.CategoryType == CategoryType.Model:
		if c.SubCategories.Size > 0 or c.CanAddSubcategory:
			model_cat.append(c.Name)
			mod_c.append(c)
			
sortlist = sorted(model_cat)

res = forms.SelectFromList.show(
        {'All': sortlist},
        title='Categories Selector',
        group_selector_title='All:',
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



def filcategorieslevinst(document,Category,Level): #Filtra tutti gli elementi in base a una lista di categorie e al loro livello
	if isinstance(Category, list):
		categoriesInstancelevCollector = []
		for nId in categoriesId:
			categoriesInstancelevCollector.append(FilteredElementCollector(document).OfCategoryId(nId).WherePasses(filterlevel).WhereElementIsNotElementType().ToElements())
		return categoriesInstancelevCollector

output = filcategorieslevinst(doc,category,level)

outputID = []

for o in output:
	for x in o:
		outputID.append(x.Id)

collection = List[ElementId](outputID)

select = uidoc.Selection.SetElementIds(collection)

print(categoriesId,bic)

#WherePasses(filterlevel)
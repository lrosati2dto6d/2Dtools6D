"""Select Elements based on Category and Type, 1 - Select one or more Categories. 2 - Select Type"""

__title__= 'Multi-Category\n and Type'
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

categories =doc.Settings.Categories

model_cat = []
model_cate = []
for c in categories:
	if c.CategoryType == CategoryType.Model:
		if "dwg"  not in c.Name and c.SubCategories.Size > 0 or c.CanAddSubcategory:
			model_cat.append(c.Name)
			model_cate.append(c)

sortlist = sorted(model_cat)

value = forms.SelectFromList.show(
        {'All Categories': sortlist},
        title='Select Categories',
        multiselect=True
    )

category =[]
namer = []

for ci in model_cate:
	namer.append(ci.Name)

for n,c in zip(namer,model_cate):
	for r in value:
		if n == r:
			category.append(c)

categoriesId = []
bic = []

for c in category:
	categoriesId.append(c.Id)
	bic.append(System.Enum.ToObject(BuiltInCategory, c.Id.IntegerValue))

fam_typ = []
#fam_type = FilteredElementCollector(doc,doc.ActiveView.Id).OfCategory(category).WhereElementIsElementType().ToElements()
for c in bic:
	fam_typ.append(FilteredElementCollector(doc).OfCategory(c).WhereElementIsElementType().ToElements())

name_type = []

for t in fam_typ:
	for c in t:
		name_type.append(c.Category.Name+' | '+c.LookupParameter('Type Name').AsString())

value_ty = forms.SelectFromList.show(
        {'All': name_type},
        title='Select Type',
        multiselect=True
    )

result_ty =[]
typ_name_result = []
for v in value_ty:
	result_ty.append(v.split(" | ",1))

for n in result_ty:
	typ_name_result.append(n[1])


selected_option = forms.CommandSwitchWindow.show(
		['YES','NO'],
		message='Select in Active View?',
)

collection = []
if selected_option == 'Yes':
	for c in bic:
		collection.append(FilteredElementCollector(doc,doc.ActiveView.Id).OfCategory(c).WhereElementIsNotElementType().ToElements())
else:
	for c in bic:
		collection.append(FilteredElementCollector(doc).OfCategory(c).WhereElementIsNotElementType().ToElements())

selectedID = []
for c in collection:
	for i in c:
		if i.Name in typ_name_result:
			selectedID.append(i.Id)


revit.get_selection().set_to(selectedID)

outputr = script.get_output()

outputr.print_md(	'# Types Selected:'.format(value_ty))

for n in (value_ty):
	outputr.print_md(	'## **{}**'.format(n))
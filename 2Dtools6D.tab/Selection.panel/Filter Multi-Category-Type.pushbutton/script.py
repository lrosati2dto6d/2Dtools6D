"""Based on Rectangular Selection, Filter elements by categories and types, 1 - Select one or more Categories. 2 - Select Type"""

__title__= 'Filter Multi-Category\n and Type'
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

output = revit.get_selection()

model_cat = []
model_cate = []
for c in categories:
	if c.CategoryType == CategoryType.Model:
		if "dwg"  not in c.Name and c.SubCategories.Size > 0 or c.CanAddSubcategory:
			model_cat.append(c.Name)
			model_cate.append(c)


cat = []
ids = []
for o in output:
	cat.append(o.Category.Name)
	ids.append(o.Id)

collectionids = List[ElementId](ids)

value = forms.SelectFromList.show(
        {'All Categories': set(cat)},
        title='Select Categories of Seleceted Elements',
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
name_bic = []
for c in category:
	categoriesId.append(c.Id)
	bic.append(System.Enum.ToObject(BuiltInCategory, c.Id.IntegerValue))
	name_bic.append(c.Name)

separator = " |---| "

fam_typ = []
for o in output:
	if o.Category.Name in name_bic:
		fam_typ.append (o.Category.Name+separator+o.Name)


value_ty = forms.SelectFromList.show(
        {'All': sorted(set(fam_typ))},
        title='Select Type of Selected Elements and Categories',
        multiselect=True
    )

result_ty =[]
typ_name_result = []
for v in value_ty:
	result_ty.append(v.split(separator,1))

for n in result_ty:
	typ_name_result.append(n[1])

select_elements = []

for e in output:
	if e.Category.Name in name_bic:
		select_elements.append(e)

selectedID = []

for c in select_elements:
	if c.Name in typ_name_result:
		selectedID.append(c.Id)

revit.get_selection().set_to(selectedID)

from pyrevit import script

outputr = script.get_output()

outputr.print_md(	'# Types Selected:'.format(value_ty))

for n in (value_ty):
	outputr.print_md(	'## **{}**'.format(n))

"""Select Elements based on Multi-Categories and Level, 1 - Select a Level. 2 - Select One or some Categories"""

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

leid = level.Id

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
        {'All': sortlist,
		'System':['Cable Trays','Ceilings','Conduits','Ducts','Floors','Pipes','Railings','Roofs','Stairs','Walls'],
		'Arch': ['Areas','Casework','Ceiling','Column','Curtain Panels','Curtain Wall Mullions','Doors','Floors','Furniture','Furniture System','Generic Models','Planting','Railings','Roofs','Rooms','Shaft Openings','Site','Specialty Equipment','Stairs','Topography','Walls','Windows'],
		"Stru": ['Structural Area Reinforcement', 'Structural Beam Systems', 'Structural Columns', 'Structural Connections', 'Structural Fabric Areas', 'Structural Fabric Reinforcement', 'Structural Foundations', 'Structural Framing', 'Structural Path Reinforcement', 'Structural Rebar', 'Structural Rebar Couplers', 'Structural Stiffeners', 'Structural Trusses'],
		'Hvac':['Air Terminals','Duct Accessories', 'Duct Fittings', 'Duct Insulations', 'Duct Linings', 'Duct Placeholders', 'Duct Systems', 'Ducts','Mechanical Equipment'],
		'Plum': ['Flex Pipes','Mechanical Equipment','Pipe Accessories', 'Pipe Fittings', 'Pipe Insulations', 'Pipe Placeholders', 'Pipes', 'Piping Systems','Plumbing Fixtures','Sprinklers'],
		'Elec': ['Cable Tray Fittings', 'Cable Trays','Communication Devices','Conduit Fittings', 'Conduits','Data Devices','Electrical Equipment', 'Electrical Fixtures','Fire Alarm Devices','Lighting Devices', 'Lighting Fixtures','Nurse Call Devices','Security Devices','Telephone Devices'],},
        title='Categories Selector',
        group_selector_title='Select Discipline',
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

mepduct = []
mepcable = []
meppipe = []
mepconduit = []
strframing = []

if "Ducts" in res:
	elemsd = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_DuctCurves).WhereElementIsNotElementType().ToElements()
	if len(elemsd) != 0:
		for e in elemsd:
			if e.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId() == leid:
				mepduct.append(e)

if "Pipes" in res:
	elemsp = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_PipeCurves).WhereElementIsNotElementType().ToElements()
	if len(elemsp) != 0:
		for e in elemsp:
			if e.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId() == leid:
				meppipe.append(e)

if "Cable Trays" in res:
	elemsc = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_CableTray).WhereElementIsNotElementType().ToElements()
	if len(elemsc) != 0:
		for e in elemsc:
			if e.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId() == leid:
				mepcable.append(e)

if "Conduits" in res:
	elemsco = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Conduit).WhereElementIsNotElementType().ToElements()
	if len(elemsco) != 0:
		for e in elemsco:
			if e.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId() == leid:
				mepconduit.append(e)

if "Structural Framing" in res:
	elemst = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFraming).WhereElementIsNotElementType().ToElements()
	if len(elemst) != 0:
		for e in elemst:
			if e.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).AsElementId() == leid:
				strframing.append(e)

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

if len(mepduct) != 0:
	if "Ducts" in res:
		for d in mepduct:
			outputID.append(d.Id)

if len(meppipe) != 0:
	if "Pipes" in res:
		for p in meppipe:
			outputID.append(p.Id)

if len(mepcable) != 0:
	if "Cable Trays" in res:
		for c in mepcable:
			outputID.append(c.Id)

if len(mepconduit) != 0:
	if "Conduits" in res:
		for co in mepconduit:
			outputID.append(co.Id)

if len(strframing) != 0:
	if "Structural Framing" in res:
		for st in strframing:
			outputID.append(st.Id)

collection = List[ElementId](outputID)

select = uidoc.Selection.SetElementIds(collection)

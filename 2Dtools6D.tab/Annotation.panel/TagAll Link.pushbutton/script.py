"""The tool TagAll elements of a category within a single chosen link_ Choose the category to tag, the view on which to insert the tags, the reference level and the family to use for the tag"""

__title__= "Link TagAll\n2.0"
__author__= "Luca Rosati"

import System
import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
import Autodesk

clr.AddReference('RevitAPIUI')


clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)

clr.AddReference("RevitServices")

from System.Collections.Generic import *

def ConvUnitsFM(number): #Feet to m
	output = number/0.3048
	return output

def ConvUnitsFMM(number): #Feet to mm
	output = number/304.8
	return output

from pyrevit import forms
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

lnks = FilteredElementCollector(doc).OfClass(RevitLinkInstance)

doclnk=[]
doclnkn=[]

for i in lnks:
	doclnk.append(i.GetLinkDocument())
	try:
		doclnkn.append(i.GetLinkDocument().Title)
	except:
		pass

if len(doclnk) == 0:
	forms.alert('There are no links in the current document, please upload at least one', exitscript=True)


str_linkname = forms.ask_for_one_item(
    doclnkn,
    default = doclnkn[0],
    prompt='Select a Link From List',
    title='Link Selector')

if str_linkname == None:
	forms.alert('Please select one Link', exitscript=True)

ldoc =[] 
linkinst = []
for i,j,ins in zip(doclnk,doclnkn,lnks):
	if str_linkname == j:
			ldoc=i
			linkinst=ins

selected_option = forms.CommandSwitchWindow.show(
    ['Active View', 'Select Plan View'],
     message='Select the View on which to add the Tags',
)

if selected_option == 'Active View':
	floorplan = doc.ActiveView
else:
	floorplanlist = forms.select_views(
		title='Select FloorPlan View',
		filterfunc=lambda x: x.ViewType == ViewType.FloorPlan)
	floorplan= floorplanlist[0]


levels = FilteredElementCollector(ldoc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()

level_ns =[]

for l in levels:
	level_ns.append(l.Name)

level_n = forms.ask_for_one_item(
    level_ns,
    default = level_ns[0],
    prompt='Select Level of Link',
    title='Link Level Selector')

for lev in levels:
	if lev.Name == level_n:
		level = lev

leid = level.Id

filterlevel = ElementLevelFilter(leid)

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
        multiselect=False
    )

category = []
namer = []

for ci in mod_c:
	namer.append(ci.Name)

for n,c in zip(namer,mod_c):
	if n == res:
		category = c

categoriesId = category.Id
bic = System.Enum.ToObject(BuiltInCategory, category.Id.IntegerValue)

tagsname = res+' Tags'
tagsname_n = res+'Tags'
tagsname_nn = res[:-1]+ 'Tag'
tagsname_nnn = res[:-1]+ ' Tags'

model_cat_ann = []
mod_c_ann = []

for c in categories:
	if c.CategoryType == CategoryType.Annotation and c.IsTagCategory:
		model_cat_ann.append(c.Name)
		mod_c_ann.append(c)
		
sortlistann = sorted(model_cat_ann)

categoryann = None
namerann = []

for cin in mod_c_ann:
	namerann.append(cin.Name)

for nann,cann in zip(namerann,mod_c_ann):
	if nann == tagsname or nann == tagsname_n or nann == tagsname_nn or nann == tagsname_nnn:
		categoryann = cann

categoriesIdann = categoryann.Id
bicann = System.Enum.ToObject(BuiltInCategory, categoryann.Id.IntegerValue)

symb = FilteredElementCollector(doc).OfCategory(bicann).WhereElementIsElementType().ToElements()

symbname=[]

if len(symb) == 0:
	forms.alert('No {} family is loaded in the project. Please load at least one and Run again the tool'.format(categoryann.Name), exitscript=True)
else:
	for s in symb:
		symbname.append(s.FamilyName + " - " + s.LookupParameter('Type Name').AsString())

tags_n = forms.ask_for_one_item(
    symbname,
    default = symbname[0],
    prompt='Select Family Tag',
    title='Family Tag Selector')

symb_selec = None

for te,tn in zip(symb,symbname):
	if tags_n == tn:
		symb_selec = te

level_filt = ElementLevelFilter(leid)

locations = []
refs = []

elem=[]

leader_option = forms.CommandSwitchWindow.show(
    ['Yes', 'No'],
     message='Do you want Add a Leader?',
)

orientation_option = forms.CommandSwitchWindow.show(
    ['Horizontal', 'Vertical'],
     message='Select an Orientation',
)

leader_bool = None

if leader_option == 'Yes':
	leader_bool = True
else:
	leader_bool = False


orient_bool = None

if orientation_option == 'Horizontal':
	orient_bool = Autodesk.Revit.DB.TagOrientation.Horizontal
else:
	orient_bool = Autodesk.Revit.DB.TagOrientation.Vertical


if bic == BuiltInCategory.OST_DuctCurves:
	elemcoll = FilteredElementCollector(ldoc).OfCategory(bic).WhereElementIsNotElementType().ToElements()
	if len(elemcoll) != 0:
		for e in elemcoll:
			if e.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId() == leid:
				elem.append(e)

elif bic == BuiltInCategory.OST_PipeCurves:
	elemcoll = FilteredElementCollector(ldoc).OfCategory(bic).WhereElementIsNotElementType().ToElements()
	if len(elemcoll) != 0:
		for e in elemcoll:
			if e.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId() == leid:
				elem.append(e)

elif bic == BuiltInCategory.OST_CableTray:
	elemcoll = FilteredElementCollector(ldoc).OfCategory(bic).WhereElementIsNotElementType().ToElements()
	if len(elemcoll) != 0:
		for e in elemcoll:
			if e.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId() == leid:
				elem.append(e)

elif bic == BuiltInCategory.OST_Conduit:
	elemcoll = FilteredElementCollector(ldoc).OfCategory(bic).WhereElementIsNotElementType().ToElements()
	if len(elemcoll) != 0:
		for e in elemcoll:
			if e.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId() == leid:
				elem.append(e)

elif bic == BuiltInCategory.OST_StructuralFraming:
	elemcoll = FilteredElementCollector(ldoc).OfCategory(bic).WhereElementIsNotElementType().ToElements()
	if len(elemcoll) != 0:
		for e in elemcoll:
			if e.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).AsElementId() == leid:
				elem.append(e)
else:
	elem = FilteredElementCollector(ldoc).OfCategory(bic).WherePasses(level_filt).WhereElementIsNotElementType().ToElements()
	if len(elem) == 0:
		elemcoll = FilteredElementCollector(ldoc).OfCategory(bic).WhereElementIsNotElementType().ToElements()
		if len(elemcoll) != 0:
			for e in elemcoll:
				if e.LookupParameter('Level').AsElementId() == leid:
					elem.append(e)


if len(elem) != 0:
	for e in elem:
		if e.Category.Name == 'Rooms':
			bb = e.get_BoundingBox(None)
			bbmax = bb.Max
			bbmin = bb.Min
			bb_center = (bbmax + bbmin)/2
			locations.append(bb_center)
			refs.append(Reference(e))
		elif e.Category.Name == 'Doors':
			if 'Open' not in e.Name:
				try:
					bb = e.get_BoundingBox(None)
					bbmax = bb.Max
					bbmin = bb.Min
					bb_center = (bbmax + bbmin)/2
					locations.append(bb_center)
					refs.append(Reference(e))
				except:
					pass
		else:
			bb = e.get_BoundingBox(None)
			bbmax = bb.Max
			bbmin = bb.Min
			bb_center = (bbmax + bbmin)/2
			locations.append(bb_center)
			refs.append(Reference(e))
else:
	forms.alert('No {} element is loaded in {}. Please load at least one and Run again the tool'.format(category.Name,str_linkname), exitscript=True)


reflink = []

for ref in refs:
	try:
		reflink.append(ref.CreateLinkReference(linkinst))
	except:
		reflink.append("Can't get this reference")

tag =[]

t = Transaction(doc,"Create Tags")

t.Start()

try:
	for e,reference,location in zip(elem,reflink,locations):
		tag.append(IndependentTag.Create(doc,symb_selec.Id,floorplan.Id,reference,leader_bool,orient_bool,location))
except:
	t.RollBack()
else:
	t.Commit()

output = script.get_output()

output.print_md(	'# Link Name: {}'.format(str_linkname))

output.print_md(	'# Category: {}'.format(category.Name))

output.print_md(	'##\tFamily Tag: {}'.format(tags_n))

output.print_md(	'##\tTags inserted: n.{}'.format(len(tag)))

for idx, elid in enumerate(tag):
		print('{}: {} - {}'.format(idx+1,elid.Name,output.linkify(elid.Id)))
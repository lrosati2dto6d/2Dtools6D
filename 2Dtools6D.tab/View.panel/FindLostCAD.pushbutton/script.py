"""Find Lost Cad"""

__title__= 'Find Lost\nCAD'
__author__= 'Luca Rosati'

import clr

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import *

clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(Revit.Elements)

clr.AddReference('RevitServices')
from System.Collections.Generic import *

from pyrevit import forms
from pyrevit import script

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument

output = script.get_output()

output.resize(1200,800)

coll_classe = FilteredElementCollector(doc).OfClass(ImportInstance).WhereElementIsNotElementType().ToElements()


if len(coll_classe) == 0:
	forms.alert('Non ci sono CAD linkati/importati in questo Modello', exitscript=True)

name_cad = []
name_cad_unique = set()
views_cad = []

for c in coll_classe:
	name_cad.append(doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM_MT).AsValueString())
	
	name_cad_unique.add(doc.GetElement(c.GetTypeId()).get_Parameter(BuiltInParameter.ELEM_CATEGORY_PARAM_MT).AsValueString())
	
	views_cad.append(doc.GetElement(c.OwnerViewId))

name_cad_x = []

for x in name_cad_unique:
	name_cad_x.append(x)

str_linkname = forms.ask_for_one_item(
    name_cad_x,
    default = name_cad_x[0],
    prompt='Select the Link CAD Name',
    title='Find Lost CAD')

if str_linkname == None:
	forms.alert('Selezionare almento un CAD linkato/importato', exitscript=True)

result = []
result_Name = []
result_type = []

for i,x in zip(name_cad,views_cad):
	if str_linkname == i:
		result.append(":heavy_multiplication_x: {} - {}".format(x.Title,output.linkify(x.Id)))
		result_Name.append(x.Name)
		result_type.append(x.Title)


output.print_md(	'# - :black_square_button: VISTE ASSOCIATE AL CAD: => "{}"'.format(str_linkname))

for n in result:
	output.print_md(	'## {} '.format(n))

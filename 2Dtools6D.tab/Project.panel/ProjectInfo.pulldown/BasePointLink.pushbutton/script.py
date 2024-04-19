"""Extract Shared Position from Links"""

__title__= 'BasePoint\nFrom Link'
__author__= 'Luca Rosati'

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

from pyrevit import DB
from pyrevit import forms
from pyrevit import script

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument

def ParaInst(element,paraname):
	if element.LookupParameter(paraname).StorageType == StorageType.Double:
		value = element.LookupParameter(paraname).AsDouble()
	elif element.LookupParameter(paraname).StorageType == StorageType.ElementId:
		value = element.LookupParameter(paraname).AsElementId()
	elif element.LookupParameter(paraname).StorageType == StorageType.String:
		value = element.LookupParameter(paraname).AsString()
	elif element.LookupParameter(paraname).StorageType == StorageType.Integer:
		value = element.LookupParameter(paraname).AsInteger()
	elif element.LookupParameter(paraname).StorageType == None:
		value = "Da Compilare"
	return value

#-------------------------------------FINESTRA OUTPUT

output = script.get_output()

output.resize(800,800)

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

b_point_list = []

for d in doclnk:
	b_point_list.append(DB.FilteredElementCollector(d).OfClass(BasePoint).WhereElementIsNotElementType().ToElements())

base_point_list = []
for bp in b_point_list:
	for b in bp:
		if b.Category.Id.IntegerValue == -2001271:
			base_point_list.append(b)

NS_values = []
EW_values = []
Z_values = []

currentDisplayUnits = doc.GetUnits().GetFormatOptions(DB.SpecTypeId.Length).GetUnitTypeId()

for bpl in base_point_list:

	X_position = DB.UnitUtils.ConvertFromInternalUnits(bpl.SharedPosition.X,currentDisplayUnits)
	Y_position = DB.UnitUtils.ConvertFromInternalUnits(bpl.SharedPosition.Y,currentDisplayUnits)
	Z_position = DB.UnitUtils.ConvertFromInternalUnits(bpl.SharedPosition.Z,currentDisplayUnits)

	NS_values.append(Y_position)
	EW_values.append(X_position)
	Z_values.append(Z_position)


"""
X_position = DB.UnitUtils.Convert(bpl.Position.X, DB.UnitTypeId.Feet, DB.UnitTypeId.Millimeters)
Y_position = DB.UnitUtils.Convert(bpl.Position.Y, DB.UnitTypeId.Feet, DB.UnitTypeId.Millimeters)
Z_position = DB.UnitUtils.Convert(bpl.Position.Z, DB.UnitTypeId.Feet, DB.UnitTypeId.Millimeters)
"""

output = script.get_output()
output.print_md("# LINK BASE POINT INFO: {} - Current Model".format(doc.Title))
output.print_md("## Current Length Unit: {} ".format(currentDisplayUnits.TypeId.split(':')[1]))

for n,y,x,z in zip (doclnkn,NS_values,EW_values,Z_values):
	output.print_md("### Link: [{}]\nN/S: {}\n - E/W: {}\n - Elevation: {}".format(n,y,x,z))
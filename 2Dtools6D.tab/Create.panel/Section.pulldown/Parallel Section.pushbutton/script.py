"""Create a Parallel Section View located near an Element Selected - Categories: Walls,Pipe,Duct,Cable Tray,Railing"""

__title__= "Parallel \nBy Element"
__author__= "Luca Rosati"

import System
import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
import Autodesk

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import ObjectType


clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager

from System.Collections.Generic import *

def ConvUnitsFM(number): #Feet to m
	output = number/0.3048
	return output

def ConvUnitsFMM(number): #Feet to mm
	output = number/304.8
	return output

from collections import defaultdict
from pyrevit import HOST_APP
from pyrevit.framework import List
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

views = FilteredElementCollector(doc).OfClass(ViewFamilyType).WhereElementIsElementType().ToElements()

units = doc.GetUnits().GetFormatOptions(UnitType.UT_Length).DisplayUnits

numbers = range(100,1000,100)

value_n = forms.ask_for_one_item(
    numbers,
    default= numbers[4],
    prompt='Select Offset Length (mm)',
    title='Create Section View')

if units == DisplayUnitType.DUT_MILLIMETERS:
	xs = ConvUnitsFMM(value_n)
elif units == DisplayUnitType.DUT_METERS:
	xs = ConvUnitsFM(value_n/1000.0)
else:
	forms.alert('Set Project Units to Millimeters or Meters ', exitscript=True)

v_names=[]
sect_views = []

for v in views:
	if v.ViewFamily.ToString() =="Section":
		v_names.append(v.LookupParameter("Type Name").AsString())
		sect_views.append(v)

value = forms.ask_for_one_item(
    v_names,
    default= v_names[0],
    prompt='Select View Type',
    title='Create Section View')

section_type = None

for i,j in zip(v_names,sect_views):
	if value == i:
		section_type = j

with forms.WarningBar(title='Pick source object:'):
	sel = uidoc.Selection 
	selected = sel.PickObject(ObjectType.Element)

unwr = doc.GetElement(selected)

if unwr.Category.Name not in ["Walls","Ducts","Pipes","Cable Trays","Railings"]:
	forms.alert('The selected object must necessarily be a Wall, a Pipe or a Duct or a Railing', exitscript=True)

bb = unwr.get_BoundingBox(None)

if unwr.Category.Name == "Walls": 
	h = unwr.LookupParameter("Unconnected Height").AsDouble()
elif unwr.Category.Name == "Railings":
	t = doc.GetElement(unwr.LookupParameter("Type").AsElementId())
	h = t.LookupParameter("Railing Height").AsDouble()


try:
	if unwr.Category.Name == "Railings":
		path = unwr.GetPath()
		if len(path)== 1:
			line = path[0]
		else:
			npaths = range(0,len(path))
			
			valuep = forms.ask_for_one_item(
			npaths,
			default= npaths[0],
			prompt='Select Railing Segment Number',
			title='Number of Path Segments {}'.format(len(path)))
			
			line = path[valuep]
	else:
		line = unwr.Location.Curve

	p = line.GetEndPoint(0);
	q = line.GetEndPoint(1);
	v = q - p

	minZ = bb.Min.Z
	maxZ = bb.Max.Z

	w = v.GetLength()
	offset = w 

	min = XYZ( -w/2-xs, -xs, -xs )
	if unwr.Category.Name in ["Walls","Railings"]:
		max = XYZ( w/2+xs, h+xs, +xs )
	else:
		max = XYZ( w/2+xs, +xs, +xs )

	midpoint = p + 0.5 * v
	walldir = v.Normalize()
	up = XYZ.BasisZ
	viewdir = walldir.CrossProduct(up)

	t = Transform.Identity
	t.Origin = midpoint
	t.BasisX = walldir
	t.BasisY = up
	t.BasisZ = viewdir

	sectionBox = BoundingBoxXYZ()
	sectionBox.Transform = t
	sectionBox.Min = min
	sectionBox.Max = max

	t = Transaction(doc,"Create Section")

	t.Start()
	try:
		section = ViewSection.CreateSection(doc,section_type.Id,sectionBox)
	except:
		t.RollBack()
	else:
		t.Commit()
except:
	forms.alert('The selected object it must not have a slope', exitscript=True)
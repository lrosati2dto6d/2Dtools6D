"""Create a Perpendicular Section View located near an Element Selected"""

__title__= "Perpendicular\nTo Element"
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

try:
	units = doc.GetUnits().GetFormatOptions(UnitType.UT_Length).DisplayUnits
except:
	units = doc.GetUnits().GetFormatOptions(SpecTypeId.Length).GetUnitTypeId()

numbers = range(100,2100,100)

value_n = forms.ask_for_one_item(
    numbers,
    default= numbers[4],
    prompt='OFFSET LENGTH (mm)',
    title='Create Section View')

try:
	if units == DisplayUnitType.DUT_MILLIMETERS:
		xs = ConvUnitsFMM(value_n)
	elif units == DisplayUnitType.DUT_METERS:
		xs = ConvUnitsFM(value_n/1000.0)
	else:
		forms.alert('Set Project Units to Millimeters or Meters ', exitscript=True)

except:
	if units == UnitTypeId.Millimeters:
		xs = ConvUnitsFMM(value_n)
	elif units == UnitTypeId.Meters:
		xs = ConvUnitsFM(value_n/1000.0)
	else:
		forms.alert('Set Project Units to Millimeters or Meters ', exitscript=True)

v_names=[]
sect_views = []

for v in views:
	if v.ViewFamily.ToString() =="Section":
		v_names.append(v.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString())
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

"""
if unwr.Location.GetType()!= LocationCurve and unwr.Location.GetType()!= LocationPoint and unwr.Category.Name != "Railings":
	forms.alert('The selected object must necessarily be Line or Point Based', exitscript=True)
"""

bb = unwr.get_BoundingBox(None)

if unwr.Category.Id.IntegerValue == -2000011: 
	h = unwr.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble()
elif unwr.Category.Id.IntegerValue == -2000126:
	t = doc.GetElement(unwr.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsElementId())
	h = t.get_Parameter(BuiltInParameter.STAIRS_RAILING_HEIGHT).AsDouble()


if unwr.Location == None:
	point_l = unwr.GetTransform().Origin
	dirh_el = unwr.HandOrientation
	ndirh_el = dirh_el.Negate()
	ro = point_l.Add(dirh_el)
	so = point_l.Add(ndirh_el)
	
	line = Line.CreateBound(ro, so)
	
elif unwr.Category.Id.IntegerValue == -2000126:
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

elif unwr.Location.GetType()== LocationPoint:
	point_l = unwr.Location.Point
	dirh_el = unwr.HandOrientation
	ndirh_el = dirh_el.Negate()
	ro = point_l.Add(dirh_el)
	so = point_l.Add(ndirh_el)
	
	line = Line.CreateBound(ro, so)
else:
	line = unwr.Location.Curve


try:
	c_tra = line.ComputeDerivatives( 0.5, True )
	origin = c_tra.Origin
	viewdir = c_tra.BasisX.Normalize()
	up = XYZ.BasisZ
	right = up.CrossProduct( viewdir )

	t = Transform.Identity
	t.Origin = origin
	t.BasisX = right
	t.BasisY = up
	t.BasisZ = viewdir

	d = xs
	minZ = bb.Min.Z
	maxZ = bb.Max.Z
	he = maxZ-minZ

	if unwr.Location == None or unwr.Location.GetType()== LocationPoint:
		minX = bb.Min.X
		maxX = bb.Max.X

		minY = bb.Min.Y
		maxY = bb.Max.Y

		lx = maxX-minX
		ly = maxY-minY

		leng = (lx,ly)

		xs = min(leng)
		
		min = XYZ(-xs/2,-xs/4,0)
	else:
		min = XYZ(-xs , -xs,0)

	if unwr.Location == None or unwr.Location.GetType()== LocationPoint:
		max = XYZ(xs/2,he + xs/4,xs/2)
	elif unwr.Category.Id.IntegerValue in [-2000011,-2000126]:
		max = XYZ(xs,xs+he,xs)
	else:
		max = XYZ(xs,xs,xs)


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
	forms.alert('The selected object probably, it must not have a slope or is not Processable', exitscript=True)

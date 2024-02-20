"""Create a Cross Sections Views located near an Element Selected"""

__title__= "Cross\nTo Element"
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

try:
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

	p = line.GetEndPoint(0);
	q = line.GetEndPoint(1);
	v = q - p

	minZ = bb.Min.Z
	maxZ = bb.Max.Z

	w = v.GetLength()

	if unwr.Location == None or unwr.Location.GetType()== LocationPoint:
		minX = bb.Min.X
		maxX = bb.Max.X

		minY = bb.Min.Y
		maxY = bb.Max.Y

		lx = maxX-minX
		ly = maxY-minY

		leng = (lx,ly)

		xs = max(leng)
		
		he = maxZ-minZ
		
		min0 = XYZ( -w/2-xs/2, -xs/4,0)
	else:
		min0 = XYZ( -w/2-xs, -xs,0)

	if unwr.Location == None or unwr.Location.GetType()== LocationPoint:
		max0 = XYZ( w/2+xs/2, he+xs/4, +xs/2 )
	elif unwr.Category.Id.IntegerValue in [-2000011,-2000126]:
		max0 = XYZ( w/2+xs, h+xs, +xs )

	else:
		max0 = XYZ( w/2+xs, +xs, +xs )

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
	sectionBox.Min = min0
	sectionBox.Max = max0

	t0 = Transaction(doc,"Create Section parallel")

	t0.Start()

	try:
		section = ViewSection.CreateSection(doc,section_type.Id,sectionBox)
	except:
		t0.RollBack()
	else:
		t0.Commit()

	c_tra1 = line.ComputeDerivatives( 0.5, True )
	origin1 = c_tra1.Origin
	viewdir1 = c_tra1.BasisX.Normalize()
	up1 = XYZ.BasisZ
	right1 = up1.CrossProduct( viewdir1 )

	t1 = Transform.Identity
	t1.Origin = origin1
	t1.BasisX = right1
	t1.BasisY = up1
	t1.BasisZ = viewdir1

	d = xs
	minZ1 = bb.Min.Z
	maxZ1 = bb.Max.Z
	he = maxZ-minZ

	if unwr.Location == None or unwr.Location.GetType()== LocationPoint:
		minX = bb.Min.X
		maxX = bb.Max.X

		minY = bb.Min.Y
		maxY = bb.Max.Y

		lx = maxX-minX
		ly = maxY-minY

		leng = (lx,ly)
		xs = max(leng)
		
		min1 = XYZ(-xs/2,-xs/4,0)
	else:
		min1 = XYZ(-xs , -xs,0)

	if unwr.Location == None or unwr.Location.GetType()== LocationPoint:
		max1 = XYZ(xs/2,he + xs/4,xs/2)
	elif unwr.Category.Id.IntegerValue in [-2000011,-2000126]:
		max1 = XYZ(xs,xs+he,xs)

	else:
		max1 = XYZ(xs,xs,xs)


	sectionBox1 = BoundingBoxXYZ()
	sectionBox1.Transform = t1
	sectionBox1.Min = min1
	sectionBox1.Max = max1

	t2 = Transaction(doc,"Create Section perpendicular")

	t2.Start()
	try:
		section = ViewSection.CreateSection(doc,section_type.Id,sectionBox1)
	except:
		t2.RollBack()
	else:
		t2.Commit()


except:
	forms.alert('The selected object probably, it must not have a slope or is not Processable', exitscript=True)
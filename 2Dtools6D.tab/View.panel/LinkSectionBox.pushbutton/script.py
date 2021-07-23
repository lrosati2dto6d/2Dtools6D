"""Create a Section Box from Linked Element in 3D View"""

__title__= "Linked\nSection Box"
__author__= "Luca Rosati"

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

clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(Revit.Elements)


clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

from collections import defaultdict
from pyrevit import HOST_APP
from pyrevit.framework import List
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument


doc_v = FilteredElementCollector(doc).OfClass(View)
th3d = []
th3dn = []
view3 = None


for d in doc_v:
	if d.IsTemplate == False and d.ViewType == ViewType.ThreeD:
		th3d.append(d)
		th3dn.append(d.Name)

if len(th3dn) != 0:
	value_n = forms.ask_for_one_item(
		th3dn,
		default= None,
		prompt='If the field is left blank, then a new 3D view will be created\nselect a view from the list to use the section box in that',
		title='Select One 3d View')

	for v,n in zip (th3d,th3dn):
		if value_n == n:
			view3 = v

if view3 == None:
	doc_vtype = FilteredElementCollector(doc).OfClass(ViewFamilyType)
	
	nam = [vt.FamilyName for vt in doc_vtype]
	
	for n,vt in zip(nam,doc_vtype):
		if n == "3D View":
			vtid = vt.Id
	
	if doc.ActiveView.ViewType.ToString() != "ThreeD":
		tv = Transaction(doc,"Create 3d View")
		tv.Start()
		
		try:
		
			isoView = View3D.CreateIsometric(doc,vtid)
		
		except:
			tv.RollBack()
		else:
			tv.Commit()
	else:
		isoView = doc.ActiveView
else:
	isoView = view3


sel1 = uidoc.Selection
ot = Selection.ObjectType.LinkedElement
el_ref = sel1.PickObject(ot, "Pick a linked element.")
linkInst = doc.GetElement(el_ref.ElementId)
linkDoc = linkInst.GetLinkDocument()
linkEl = linkDoc.GetElement(el_ref.LinkedElementId)

bb = linkEl.get_BoundingBox(None)

t = Transaction(doc,"Create Section Box")

t.Start()
try:
	isoView.SetSectionBox(bb)
except:
	t.RollBack()
else:
	t.Commit()

uidoc.ActiveView = isoView
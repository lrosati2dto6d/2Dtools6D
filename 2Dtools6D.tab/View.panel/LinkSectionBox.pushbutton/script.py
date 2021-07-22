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

try:
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
		if doc.ActiveView.ViewType.ToString() == "ThreeD":
			doc.ActiveView.SetSectionBox(bb)
		else:
			forms.alert("Select Element only from a 3D View", exitscript=True)

	except:
		t.RollBack()
	else:
		t.Commit()
except:
	forms.alert("There aren't Revit Links Loaded", exitscript=True)
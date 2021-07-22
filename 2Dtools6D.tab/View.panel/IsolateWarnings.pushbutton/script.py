"""Isolate Warning Elements in a New 3D View"""

__title__= 'Isolate\nWarnings'
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

val_ask = ["Yes","No"]

value_n = forms.ask_for_one_item(
    val_ask,
    default= val_ask[0],
    prompt='Select Answer',
    title='Create Isolate Warnings 3D View?')

if value_n == "Yes":
	doc_warn = doc.GetWarnings()
	if len(doc_warn) != 0:
		warn_ele = []
		uniq = []

		for w in doc_warn:
			warn_ele.append(w.GetFailingElements())
		
		for l in warn_ele:
			for e in l:
				if e not in uniq:
					uniq.append(e)

		doc_vtype = FilteredElementCollector(doc).OfClass(ViewFamilyType)

		nam = [vt.FamilyName for vt in doc_vtype]

		for n,vt in zip(nam,doc_vtype):
			if n == "3D View":
				vtid = vt.Id

		t = Transaction(doc,"Create Isolate Warnings View")

		t.Start()
		try:
			isoView = View3D.CreateIsometric(doc,vtid)
			isoView.IsolateElementsTemporary(List[ElementId](uniq))
			
			doc.Regenerate()
			
			isoView.ConvertTemporaryHideIsolateToPermanent()
		except:
			t.RollBack()
		else:
			t.Commit()
	else:
		forms.alert('There are no Warnings to isolate ... Good job!', exitscript=True)
else:
	forms.alert('You have chosen not to create the view for the Warnings', exitscript=True)
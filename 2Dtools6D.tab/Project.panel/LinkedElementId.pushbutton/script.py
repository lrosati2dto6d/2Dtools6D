"""Gets Id of linked Element/s"""

__title__= "Linked\nElement Id"

import clr
import sys
sys.path.append('C:\Program Files (x86)\IronPython 2.7\Lib')

import Autodesk

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
from pyrevit import output

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument
Output = output.get_output()

doc_v = FilteredElementCollector(doc).OfClass(View)


sel1 = uidoc.Selection
ot = Selection.ObjectType.LinkedElement
el_ref = sel1.PickObjects(ot, "Pick a linked element.")
Output.print_md("##Linked Elements info:")
Output.insert_divider()
for item in el_ref:
    
    linkInst = doc.GetElement(item.ElementId)
    linkDoc = linkInst.GetLinkDocument()
    linkEl = linkDoc.GetElement(item.LinkedElementId)
    if linkEl.GetType() == Autodesk.Revit.DB.FamilyInstance:
        Output.print_md("###Element ID : {} of model : {}".format(linkEl.Id,linkDoc.Title))
        Output.print_md("####Family Name : {}".format(linkEl.Symbol.FamilyName))
        Output.print_md("####Type : {}".format(linkEl.LookupParameter("Type").AsValueString()))
        Output.insert_divider()
    else:
        Output.print_md("###Element ID : {} of model : {}".format(linkEl.Id,linkDoc.Title))  
        Output.print_md("####Family Type : {}".format(linkEl.LookupParameter("Type").AsValueString()))
        Output.insert_divider()

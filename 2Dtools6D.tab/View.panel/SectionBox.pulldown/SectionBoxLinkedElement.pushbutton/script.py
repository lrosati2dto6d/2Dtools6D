"""Creates a bounding box around linked element by ID"""

__title__= "Section Box On\nLinked Element by ID"

import clr
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import *

clr.AddReference('RevitAPI')
import Autodesk
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from pyrevit import forms

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument

uiviews = uidoc.GetOpenUIViews()
view = doc.ActiveView
uiview = [x for x in uiviews if x.ViewId == view.Id][0]

## Link List ##
links = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_RvtLinks).WhereElementIsNotElementType().ToElements()
linkString=[]
for link in links:
    try:
        linkString.append(link.GetLinkDocument().Title)
    except:
        linkString.append("<Unloaded Link>")
linkOfRoom = forms.SelectFromList.show(context = linkString,title = "Seleziona il link delle room", width  = 500, height = 500)
linkInstance = links[linkString.index(linkOfRoom)]
######
elemId = int(forms.ask_for_string(prompt = "Insert Linked Element ID", title = "Linked Element Id"))
selection = linkInstance.GetLinkDocument().GetElement(ElementId(elemId))

doc_v = FilteredElementCollector(doc).OfClass(View)
th3d = []
th3dn = []
view3 = None


for d in doc_v:
	if d.IsTemplate == False and d.ViewType == ViewType.ThreeD:
		th3d.append(d)
		th3dn.append(d.Name)

value_n = (forms.select_views(title = "Select One View", button_name= "Select",width=500,multiple=False)).Id

for v,n in zip (th3d,th3dn):
	if value_n == n:
		view3 = v
		break

isoView = doc.GetElement(value_n)

bb = selection.get_BoundingBox(None)

t = Transaction(doc,"Section Box On Id")

t.Start()
try:
	isoView.SetSectionBox(bb)
except:
	t.RollBack()
else:
	t.Commit()

uidoc.ActiveView = isoView

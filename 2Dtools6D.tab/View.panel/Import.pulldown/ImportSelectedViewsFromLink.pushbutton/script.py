"""Select Views from link and copy to current document"""

__title__= 'Import Views\nFrom Link'
__author__= 'Luca Rosati'

import System
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
import RevitServices
from RevitServices.Persistence import DocumentManager

from collections import defaultdict
from pyrevit import revit, DB
from pyrevit import HOST_APP
from pyrevit.framework import List
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument

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


value = forms.SelectFromList.show(
        {'All Links Loaded': doclnkn},
        title='Select Link',
        multiselect=False
    )

if value == None:
	forms.alert('There are no links selected, please select at least one', exitscript=True)

ldoc = None 
linkinst = None
for i,j,ins in zip(doclnk,doclnkn,lnks):
	if j == value:
		ldoc=i
		linkinst=ins

lnk_v = DB.FilteredElementCollector(ldoc).OfCategory(DB.BuiltInCategory.OST_Views).WhereElementIsNotElementType()
if not lnk_v:
	forms.alert('There are no Views in Link, please select at least one', exitscript=True)


lnk_v = [v for v in lnk_v if not v.IsTemplate and v.ViewType!= ViewType.Elevation]


lnk_v.sort(key=lambda v: v.Name)

selected = forms.SelectFromList.show(lnk_v, button_name='Select Views', name_attr='Name', multiselect=True)
if not selected:
	forms.alert('Please select at least one View', exitscript=True)

selected_id = [v.Id for v in selected]
selected_id = List[DB.ElementId](selected_id)

try:
    with revit.Transaction('Copy views for Link'):
        DB.ElementTransformUtils.CopyElements(ldoc, selected_id, doc, None, None)
except:
	pass



output = script.get_output()


output.print_md(	'#Link ----> [{}]'.format(ldoc.Title))

output.print_md(	'#List of correctly copied views')

for n in (selected):
	output.print_md(	'##[{}]'.format(n.Title))
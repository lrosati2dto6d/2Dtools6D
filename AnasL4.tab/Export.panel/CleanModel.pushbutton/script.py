"""Clean Modelli per la Consegna ANAS L-4"""

__title__= 'Clean Delivery\nAnas L-4'
__author__= 'Luca Rosati'

import clr
import System

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import *

clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(Revit.Elements)

clr.AddReference('RevitServices')
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
from System.Collections.Generic import *


from Autodesk.Revit.UI import RevitCommandId
from Autodesk.Revit.UI import UIApplication
from Autodesk.Revit.UI import ExternalCommandData

from collections import defaultdict
from pyrevit import HOST_APP
from pyrevit.framework import List
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script
from pyrevit.framework import Emojis

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument
uiapp = __revit__

#-------------------------------------FINESTRA OUTPUT

output = script.get_output()

output.resize(600,200)


d_wset = FilteredWorksetCollector(doc)
wset_name = []

for c in d_wset:
	wset_k = c.Kind
	if wset_k == WorksetKind.UserWorkset:
		wset_name.append(c.Name)


if len(wset_name) > 1:
	forms.alert('Risultano presenti dei Workset, Prima di procedere assicurarsi di Dissociare il Modello Ignorando i Workset', exitscript=True)

if not forms.alert('Prima di Procedere Assicurarsi che il Modello sia salvato e Dissociato sul proprio PC locale'
                   ' - Clicca OK per Confermare', cancel=True):
    script.exit()

#-------------------------------------EXP.VIEWS

pvp = ParameterValueProvider(ElementId(-1007400))
fng = FilterStringLess()
ruleValue = 'Informazioni di progetto'
fRule = FilterStringRule(pvp,fng,ruleValue,True)

filter = ElementParameterFilter(fRule)

del_viewsheet_coll = FilteredElementCollector(doc).OfClass(ViewSheet).WhereElementIsNotElementType().ToElements()

sheet_del = []
startingview = None

for d in del_viewsheet_coll:
	if 'Informazioni di progetto' not in d.Name:
		sheet_del.append(d)
	else:
		startingview = d

if startingview.Name != doc.ActiveView.Name:
	forms.alert('Per Attivare lo Script aprire la Vista Iniziale XXX - Informazioni di progetto', exitscript=True)
else:
	pass


del_views_coll = FilteredElementCollector(doc).OfClass(View).WhereElementIsNotElementType().ToElements()
view_del = []

for vdel in del_views_coll:
	if vdel.HasDetailLevel():
		view_del.append(vdel)

todelete = []
filtapply = None

del_filt_coll = FilteredElementCollector(doc).OfClass(ParameterFilterElement).WhereElementIsNotElementType().ToElements()

for vd in del_filt_coll:
	if "Elementi Nascosti" not in vd.Name:
		todelete.append(vd)
	else:
		filtapply = vd


if filtapply == None:
	forms.alert('Inserire il Filtro 2Dto6D_Anas_Elementi Nascosti prima di proseguire', exitscript=True)

del_sched_coll = FilteredElementCollector(doc).OfClass(ViewSchedule).WhereElementIsNotElementType().ToElements()

del_riq_coll = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_VolumeOfInterest).WhereElementIsNotElementType().ToElements()

t = Transaction(doc,"unpin")
t.Start()

for r in del_riq_coll:
	if r.Pinned == True:
		r.Pinned = False
t.Commit()

sheet_del.extend(del_riq_coll)
sheet_del.extend(view_del)
sheet_del.extend(del_sched_coll)
todelete.extend(sheet_del)


t = Transaction(doc,"delete")

t.Start()

for vtp in todelete: 
	try:
		doc.Delete(vtp.Id)
	except:
		pass

t.Commit()

vtid = None
doc_vtype = FilteredElementCollector(doc).OfClass(ViewFamilyType)

nam = [vt.FamilyName for vt in doc_vtype]

for n,vt in zip(nam,doc_vtype):
	if n == "Vista 3D":
		vtid = vt

d_phases = doc.Phases
fasi_list = []
phases_name = []
fase = None

for p in d_phases:
	fasi_list.append(p)
	phases_name.append(p.Name)


for p in fasi_list:
	p_name = p.Name
	if p_name == "Nuova costruzione":
		fase = p

d_phases_Fil = FilteredElementCollector(doc).OfClass(PhaseFilter)
fasef = None
for pf in d_phases_Fil:
	pf_name = pf.Name
	if pf_name == "Mostra completo":
		fasef = pf


t = Transaction(doc,"Create 3d View")

t.Start()
isoView = View3D.CreateIsometric(doc,vtid.Id)

doc.Regenerate()

isoView.AddFilter(filtapply.Id)

doc.Regenerate()

isoView.SetFilterVisibility(filtapply.Id,False)
isoView.DetailLevel = ViewDetailLevel.Fine
isoView.DisplayStyle = DisplayStyle.FlatColors

doc.Regenerate()

isoView.LookupParameter("Fase").Set(fase.Id)
isoView.LookupParameter("Filtro delle fasi").Set(fasef.Id)

doc.Regenerate()

'''
CmndID = RevitCommandId.LookupCommandId('ID_PURGE_UNUSED')
CmId = CmndID.Id
uiapp.PostCommand(CmndID)
'''

t.Commit()


output.print_md(	'# - MODELLO PULITO CORRETTAMENTE ---> :white_heavy_check_mark:')
output.print_md(	'-----------------------')
output.print_md(	'## - Prima di Salvare eseguire il Purge e Cambiare il Browser Organization :red_circle:')
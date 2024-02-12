"""Creates a table where all phases associated with links in the project are displayed"""

__title__= 'Link-Phase\nInfo'
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

if forms.check_workshared() == False:
   script.exit()

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

phase_par = []

for l in lnks:
	phase_par.append(doc.GetElement(l.GetTypeId()).GetPhaseMap())

ph_key = []
ph_val = []

for ph in phase_par:
	yu = []
	yi = []
	for p in ph:
		yu.append(doc.GetElement(p.Value).Name)
		yi.append(doc.GetElement(p.Key).Name)
	ph_val.append(yu)
	ph_key.append(yi)


output = script.get_output()

output.print_md(	'#List of Link-PhaseMap')

for p,pk,pv in zip (doclnkn,ph_key,ph_val):
	output.print_md(	'#{}'.format(p))
	for k,v in zip(pk,pv):
		if k != v:
			output.print_md(	':cross_mark: {} ----> {}'.format(k,v))
		else:
			output.print_md(	':white_heavy_check_mark: {} ----> {}'.format(k,v))
"""Overwrite Filters Overrides between view templates"""

__title__= "Override Filters\nTransfer"
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

from collections import defaultdict
from pyrevit import HOST_APP
from pyrevit.framework import List
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

viewsFilter = ElementCategoryFilter(BuiltInCategory.OST_Views)

viewTemplatesname = []
viewtemplates =[]

viewsCollector = FilteredElementCollector(doc).WherePasses(viewsFilter)
for view in viewsCollector:
	if view.IsTemplate == True:
		viewTemplatesname.append(view.Name)
		viewtemplates.append(view)


MTemplate = forms.SelectFromList.show(
        {'All ViewTemplates': viewTemplatesname},
        title='Select a Master ViewTemplate',
        multiselect=False
    )

for i,x in zip(viewtemplates,viewTemplatesname):
	if MTemplate == x:
		MTemp_ele = i

MTemp_filt = MTemp_ele.GetFilters()

fil_id=[]
fil_name=[]

for f in MTemp_filt:
	fil_id.append(doc.GetElement(f))

for f in fil_id:
	fil_name.append(f.Name)

num = len(MTemp_filt)

targ_filters = forms.SelectFromList.show(
        {'Filters of Template Master': fil_name},
        title='Select Filters to Transfer',
        multiselect = True
    )

filt_targ_ele = []

for i,x in zip(fil_id,fil_name):
	for r in targ_filters:
		if x == r:
			filt_targ_ele.append(i)

fil_over =[]
fil_vis = []

for f in filt_targ_ele:
	fil_over.append(MTemp_ele.GetFilterOverrides(f.Id))
	fil_vis.append(MTemp_ele.GetFilterVisibility(f.Id))


vtemp_targ_ele = []
viewTemplates2 = []

viewsCollector2 = FilteredElementCollector(doc).WherePasses(viewsFilter)

for v in viewsCollector2:
	if v.IsTemplate == True and v.Name != MTemplate:
		viewTemplates2.append(v.Name)
		vtemp_targ_ele.append(v)
		
sortvtemp = sorted(viewTemplates2)


targ_templates = forms.SelectFromList.show(
        {'ViewTemplates to Apply Filters Overrides': sortvtemp},
        title='Select ViewTemplates',
        multiselect = True
    )

templates_ele_target = []

FILT_TARG = []

for i,x in zip(vtemp_targ_ele,viewTemplates2):
	for r in targ_templates:
		if x == r:
			templates_ele_target.append(i)

result =[]
noresult =[]
result_add = []

t = Transaction(doc,"change fil")

t.Start()
filt_targ_name_list=[]

for fm,fg,fv in zip(filt_targ_ele,fil_over,fil_vis):
	for vt in templates_ele_target:
		flist = vt.GetFilters()
		for fl in flist:
			filt_targ_name_list.append(doc.GetElement(fl).Name)
			if doc.GetElement(fl).Name == fm.Name and vt.IsFilterApplied(fm.Id) == True:
				visibility = fv
				override = OverrideGraphicSettings()
				override.SetCutBackgroundPatternColor(fg.CutBackgroundPatternColor)
				override.SetCutBackgroundPatternId(fg.CutBackgroundPatternId)
				override.SetCutBackgroundPatternVisible(fg.IsCutBackgroundPatternVisible)
				override.SetCutForegroundPatternColor(fg.CutForegroundPatternColor)
				override.SetCutForegroundPatternId(fg.CutForegroundPatternId)
				override.SetCutForegroundPatternVisible(fg.IsCutForegroundPatternVisible)
				override.SetCutLineColor(fg.CutLineColor)
				override.SetCutLinePatternId(fg.CutLinePatternId)
				override.SetCutLineWeight(fg.CutLineWeight)
				override.SetDetailLevel(fg.DetailLevel)
				override.SetHalftone(fg.Halftone)
				override.SetProjectionLineColor(fg.ProjectionLineColor)
				override.SetProjectionLinePatternId(fg.ProjectionLinePatternId)
				override.SetProjectionLineWeight(fg.ProjectionLineWeight)
				override.SetSurfaceBackgroundPatternColor(fg.SurfaceBackgroundPatternColor)
				override.SetSurfaceBackgroundPatternId(fg.SurfaceBackgroundPatternId)
				override.SetSurfaceBackgroundPatternVisible(fg.IsSurfaceBackgroundPatternVisible)
				override.SetSurfaceForegroundPatternColor(fg.SurfaceForegroundPatternColor)
				override.SetSurfaceForegroundPatternId(fg.SurfaceForegroundPatternId)
				override.SetSurfaceForegroundPatternVisible(fg.IsSurfaceForegroundPatternVisible)
				override.SetSurfaceTransparency(fg.Transparency)
				vt.SetFilterOverrides(fl,override)
				vt.SetFilterVisibility(fl,visibility)
			else:
				try:
					vt.AddFilter(fm.Id)
					noresult.append(fm.Name)
				except:
					pass
		if len(flist) == 0:
			try:
				vt.AddFilter(fm.Id)
				result_add.append(fm.Name)
			except:
				pass

t.Commit()

t = Transaction(doc,"change fil2")

t.Start()
filt_targ_name_list=[]

for fm,fg,fv in zip(filt_targ_ele,fil_over,fil_vis):
	for vt in templates_ele_target:
		flist = vt.GetFilters()
		for fl in flist:
			filt_targ_name_list.append(doc.GetElement(fl).Name)
			if doc.GetElement(fl).Name == fm.Name and vt.IsFilterApplied(fm.Id) == True:
				visibility = fv
				override = OverrideGraphicSettings()
				override.SetCutBackgroundPatternColor(fg.CutBackgroundPatternColor)
				override.SetCutBackgroundPatternId(fg.CutBackgroundPatternId)
				override.SetCutBackgroundPatternVisible(fg.IsCutBackgroundPatternVisible)
				override.SetCutForegroundPatternColor(fg.CutForegroundPatternColor)
				override.SetCutForegroundPatternId(fg.CutForegroundPatternId)
				override.SetCutForegroundPatternVisible(fg.IsCutForegroundPatternVisible)
				override.SetCutLineColor(fg.CutLineColor)
				override.SetCutLinePatternId(fg.CutLinePatternId)
				override.SetCutLineWeight(fg.CutLineWeight)
				override.SetDetailLevel(fg.DetailLevel)
				override.SetHalftone(fg.Halftone)
				override.SetProjectionLineColor(fg.ProjectionLineColor)
				override.SetProjectionLinePatternId(fg.ProjectionLinePatternId)
				override.SetProjectionLineWeight(fg.ProjectionLineWeight)
				override.SetSurfaceBackgroundPatternColor(fg.SurfaceBackgroundPatternColor)
				override.SetSurfaceBackgroundPatternId(fg.SurfaceBackgroundPatternId)
				override.SetSurfaceBackgroundPatternVisible(fg.IsSurfaceBackgroundPatternVisible)
				override.SetSurfaceForegroundPatternColor(fg.SurfaceForegroundPatternColor)
				override.SetSurfaceForegroundPatternId(fg.SurfaceForegroundPatternId)
				override.SetSurfaceForegroundPatternVisible(fg.IsSurfaceForegroundPatternVisible)
				override.SetSurfaceTransparency(fg.Transparency)
				vt.SetFilterOverrides(fl,override)
				vt.SetFilterVisibility(fl,visibility)
				result.append(fm.Name)

			else:
				try:
					vt.AddFilter(fm.Id)
				except:
					pass

t.Commit()

output = script.get_output()
'''
# Function to display result to user
def printMessage(resultList, failedList, message, messageWarning):
	if len(resultList) != 0:
		print(message)
		print("\n".join(resultList))
	if len(failedList) != 0:
		print(messageWarning)
		print("\n".join(resultList))

# Print message
printMessage(viewsSuccess, viewsFail, "The following view's view template have been changed:",
			"View templates failed to apply:")
'''

output.print_md('# - Project: {}'.format(doc.Title))

output.print_md('-----------------------------')

output.print_md('# ViewTemplate Master: {}'.format(MTemplate))

output.print_md('-----------------------------')

output.print_md('# ViewTemplates Target:')

for vT in targ_templates:
	output.print_md('## {}'.format(vT))

output.print_md('-----------------------------')

output.print_md('# Filters Target Override:')

for fil_ovv in set(result):
	output.print_md('## {}'.format(fil_ovv))

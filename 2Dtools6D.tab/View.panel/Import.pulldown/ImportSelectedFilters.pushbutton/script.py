# -*- coding: utf-8 -*-
"""Import Selected Filters from a Open Revit Project.
"""
__title__ = 'Import Selected\nFilters'
__author__ = "Luca Rosati"

# Import commom language runtime
import clr

# Import C# List
from System.Collections.Generic import List

# Import Revit DB
from Autodesk.Revit.DB import *

# Import pyRevit forms
from pyrevit import forms
from pyrevit import script

# Store current document to variable
app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Select opened documents to transfer View Templates
selProject = forms.select_open_docs(title="Select project/s to transfer View Templates", button_name='OK', width=500, multiple=True, filterfunc=None)

# Filter Views and Filters

viewsFilter = ElementCategoryFilter(BuiltInCategory.OST_Views)
filterFilter = ElementClassFilter(ParameterFilterElement)


# Function to retrieve Filters
def retrieveFI(docList, currentDoc):
	storeDict= {}
	if isinstance(docList, list):
		for pro in docList:
			filtersCollector = FilteredElementCollector(pro).WherePasses(filterFilter)
			for filter in filtersCollector:
				storeDict[filter.Name + ' - ' + pro.Title] = filter
	elif currentDoc == True:
		filtersCollector = FilteredElementCollector(docList).WherePasses(filterFilter)
		for filter in filtersCollector:
			storeDict[filter.Name] = filter
	else:
		filtersCollector = FilteredElementCollector(docList).WherePasses(filterFilter)
		for filter in filtersCollector:
			storeDict[filter.Name+ ' - ' + pro.Title] = filter
	return storeDict


# Retrieve all filters from selected docs
filtersdocs = retrieveFI(selProject, False)

# Display select view templates form
vFilters = forms.SelectFromList.show(filtersdocs.keys(), "Select one or more Filters to transfer to the Current Project", 600, 300, multiselect=True)

# Collect all View Templates in the current document
docFilters = retrieveFI(doc, True)

# Remove view templates with same name
vFiltersNoProj = []
for vF in vFilters:
	for p in selProject:
		if p.Title in vF:
			vFiltersNoProj.append(vF.replace(p.Title, ""))

# Remove duplicated view templates from selection
newVFilters = []
uniqueFilters = []
for vNoProj, vT in zip(vFiltersNoProj, vFilters):
	if vNoProj not in uniqueFilters:
		uniqueFilters.append(vNoProj)
		newVFilters.append(vT)

vFilters = newVFilters

# Collect all views from the current document
docViewsCollector = FilteredElementCollector(doc).WherePasses(viewsFilter)


# Transform object
transIdent = Transform.Identity
copyPasteOpt = CopyPasteOptions()

# Create single transaction and start it
t = Transaction(doc, "Copy Filters")
t.Start()

try:
	for vF in vFilters:
		vFId = List[ElementId]()
		for pro in selProject:
			if pro.Title in vF:
				vFId.Add(filtersdocs[vF].Id)
				# Check if view template is used in current doc
				if vF.replace(" - " + pro.Title, "") not in docFilters.keys():
					# If not, copy the selected View Template to current project
					ElementTransformUtils.CopyElements(pro, vFId, doc, transIdent, copyPasteOpt)

except:
	t.RollBack()
else:
	t.Commit()

output = script.get_output()

output.print_md('# Projects Used for Transfer:')
for pro in selProject:
	output.print_md('## - {}'.format(pro.Title))

output.print_md('-----------------------------')

output.print_md('# Filters Transferred:')
for vF in vFilters:
	output.print_md('## - {0}'.format(vF))

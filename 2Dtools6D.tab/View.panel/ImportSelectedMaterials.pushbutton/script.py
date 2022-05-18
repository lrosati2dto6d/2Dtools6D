# -*- coding: utf-8 -*-
"""Import Selected Materials from a Open Revit Project.
"""
__title__ = 'Import Selected\nMaterials'
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

# Filter Views and Materials

viewsFilter = ElementCategoryFilter(BuiltInCategory.OST_Views)
MaterialFilter = ElementClassFilter(Material)

# Function to retrieve Materials
def retrieveFI(docList, currentDoc):
	storeDict= {}
	if isinstance(docList, list):
		for pro in docList:
			filtersCollector = FilteredElementCollector(pro).WherePasses(MaterialFilter)
			for filter in filtersCollector:
				storeDict[filter.Name + ' - ' + pro.Title] = filter
	elif currentDoc == True:
		filtersCollector = FilteredElementCollector(docList).WherePasses(MaterialFilter)
		for filter in filtersCollector:
			storeDict[filter.Name] = filter
	else:
		filtersCollector = FilteredElementCollector(docList).WherePasses(MaterialFilter)
		for filter in filtersCollector:
			storeDict[filter.Name+ ' - ' + pro.Title] = filter
	return storeDict

# Retrieve all filters from selected docs
materialsdocs = retrieveFI(selProject, False)

# Display select view templates form
vMaterials = forms.SelectFromList.show(materialsdocs.keys(), "Select one or more Materials to transfer to the Current Project", 600, 300, multiselect=True)

# Collect all View Templates in the current document
docMaterials = retrieveFI(doc, True)

# Remove Materials with same name
vMaterialsNoProj = []
for vF in vMaterials:
	for p in selProject:
		if p.Title in vF:
			vMaterialsNoProj.append(vF.replace(p.Title, ""))

# Remove duplicated view templates from selection
newvMaterials = []
uniqueMaterials = []
for vNoProj, vT in zip(vMaterialsNoProj, vMaterials):
	if vNoProj not in uniqueMaterials:
		uniqueMaterials.append(vNoProj)
		newvMaterials.append(vT)

vMaterials = newvMaterials

# Collect all views from the current document
docViewsCollector = FilteredElementCollector(doc).WherePasses(viewsFilter)

# Transform object
transIdent = Transform.Identity
copyPasteOpt = CopyPasteOptions()

# Create single transaction and start it
t = Transaction(doc, "Copy Materials")
t.Start()

try:
	for vF in vMaterials:
		vFId = List[ElementId]()
		for pro in selProject:
			if pro.Title in vF:
				vFId.Add(materialsdocs[vF].Id)
				# Check if view template is used in current doc
				if vF.replace(" - " + pro.Title, "") not in docMaterials.keys():
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

output.print_md('# Materials Transferred:')
for vF in vMaterials:
	output.print_md('## - {0}'.format(vF))

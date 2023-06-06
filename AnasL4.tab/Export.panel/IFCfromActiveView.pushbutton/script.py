"""Select The property set mapping file and the destination folder of the IFC File. The script will export a 2x3 IFC format file with the mapping selected in the destination folder"""

__title__= 'Export IFC\n ActiveView'
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

from collections import defaultdict
from pyrevit import HOST_APP
from pyrevit.framework import List
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument

filepathPset = forms.pick_file(file_ext='txt')
folder=forms.pick_folder(title='Select Destination Folder', owner=None)

'''
fileversion = forms.ask_for_one_item(
        ['IFC2x3', 'IFC4 Reference', 'IFC4 Design'],
        default='IFC2x3',
        prompt='Select an IFC Version from list',
        title='Anas L-4 IFC Version'
    )
'''

fileversion = 'IFC2x3'

views_exp = forms.select_views()
names=[]

for v in views_exp:
	names.append(v.Name)

for n,v in zip(names,views_exp):

	t = Transaction(doc,'export'+n)
	t.Start()
	result = []

	options=IFCExportOptions()

	if fileversion == 'IFC4 Design':
		options.FileVersion = IFCVersion.IFC4DTV
	if fileversion == 'IFC4 Reference':
		options.FileVersion = IFCVersion.IFC4RV
	if fileversion == 'IFC2x3':
		options.FileVersion = IFCVersion.IFC2x3CV2
		
	options.WallAndColumnSplitting = False
	options.ExportBaseQuantities = False
	options.FilterViewId = v.Id
	options.AddOption("ExportInternalRevitPropertySets","false");
	options.AddOption("ExportIFCCommonPropertySets","false");	
	options.AddOption("ExportAnnotations ","false");
	options.AddOption("SpaceBoundaries ", "0");
	options.AddOption("ExportRoomsInView", "true");	
	options.AddOption("Use2DRoomBoundaryForVolume ", "false");
	options.AddOption("UseFamilyAndTypeNameForReference ", "false");
	options.AddOption("Export2DElements", "false");
	options.AddOption("ExportPartsAsBuildingElements", "true");
	options.AddOption("ExportBoundingBox", "false");
	options.AddOption("ExportSolidModelRep", "false");
	options.AddOption("ExportSchedulesAsPsets", "false");
	options.AddOption("ExportSpecificSchedules", "false");
	options.AddOption("ExportLinkedFiles", "false");
	options.AddOption("IncludeSiteElevation","true");
	options.AddOption("StoreIFCGUID", "true");
	options.AddOption("VisibleElementsOfCurrentView ", "true");
	options.AddOption("UseActiveViewGeometry", "true");
	options.AddOption("TessellationLevelOfDetail", "1");
	options.AddOption("ExportUserDefinedPsets","true");
	options.AddOption("SitePlacement", "0");
	options.AddOption("ExportUserDefinedPsetsFileName",filepathPset)
	c = doc.Export(folder, n, options)
	result.append(c)

	t.Commit()

from pyrevit import script

output = script.get_output()

for n in names:
	output.print_md(	'# **IFC NAME: {} Success'.format(n))
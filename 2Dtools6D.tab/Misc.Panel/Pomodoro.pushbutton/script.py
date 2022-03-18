# -*- coding: utf-8 -*-
"""For many people, time is an enemy. We race against the clock to finish assignments and meet deadlines"""

__title__ = 'Pomodoro\nTechnique'
__author__ = "Luca Rosati"

# Import commom language runtime
import clr

# Import C# List
from System.Collections.Generic import List

# Import Revit DB
from Autodesk.Revit.DB import FilteredElementCollector, ElementTransformUtils, BuiltInCategory, \
                            ElementId, Transform, CopyPasteOptions, Transaction, TransactionGroup, ElementCategoryFilter

# Import pyRevit forms
from pyrevit import forms
from pyrevit import script
from pyrevit.forms import ProgressBar

# Store current document to variable
app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

output = script.get_output()

max_value = 1000

with ProgressBar(title= 'POMODORO TECHNIQUE',cancellable=True,step=1) as pb:
	for counter in range(0,max_value):
		if pb.cancelled:
			forms.alert('You should not stop', exitscript=True)
		else:
			pb.update_progress(counter,max_value)

forms.alert('Good Job!!! You Deserve 5 Minutes of Coffee Break!!!', exitscript=True)
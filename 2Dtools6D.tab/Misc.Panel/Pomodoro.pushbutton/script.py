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

valueuniquitem = forms.ask_for_one_item(
    ['5 Minutes', '10 Minutes', '25 Minutes'],
    default='25 Minutes',
    prompt='Select Pomodoro Time',
    title='POMODORO TECHNIQUE')

output = script.get_output()

if valueuniquitem == '5 Minutes':
	max_value = 15000
elif valueuniquitem == '10 Minutes':
	max_value = 30000
else:
	max_value = 75000

with ProgressBar(title= 'POMODORO TECHNIQUE',cancellable=True,step=1) as pb:
	for counter in range(0,max_value):
		if pb.cancelled:
			forms.alert('You should Not stop!!! Go back to work', exitscript=True)
		else:
			pb.update_progress(counter,max_value)

forms.alert('GOOD JOB!!! You Deserve 5 Minutes of Coffee Break!!!', exitscript=True)
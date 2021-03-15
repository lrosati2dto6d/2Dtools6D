"""Get Info about Project Parameters and Category assigned"""

__title__= 'Get ProjectParameters\nInfo'
__author__= 'Luca Rosati'

import System
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
import Autodesk

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)

clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager

from pyrevit import forms


doc = __revit__.ActiveUIDocument.Document
uidoc =  __revit__.ActiveUIDocument

names = []
groups = []
pgroup = []
ptype = []
units = []
isvis = []
elements = []
guids = []
isinst = []
bics = []

iterator = doc.ParameterBindings.ForwardIterator()

while iterator.MoveNext():
	
	
	groups.append(iterator.Key.VariesAcrossGroups)
	names.append(iterator.Key.Name)
	pgroup.append(iterator.Key.ParameterGroup)
	ptype.append(iterator.Key.ParameterType)
	units.append(iterator.Key.UnitType)
	isvis.append(iterator.Key.Visible)
	
	elem = doc.GetElement(iterator.Key.Id)
	elements.append(elem)



	if elem.GetType().ToString() == 'Autodesk.Revit.DB.SharedParameterElement':
		guids.append(elem.GuidValue)
	else: guids.append(None)
	


	if iterator.Current.GetType().ToString() == 'Autodesk.Revit.DB.InstanceBinding':
		isinst.append("Instance")
	else:
		isinst.append("Type")

titoli = ["Parameter Name","IsInstance?","ParameterGroup","ParameterType","Units","GUID","Vary by ModelGroup?","IsVisible?"]

print(titoli,names,isinst,pgroup,ptype,units,guids,groups,isvis)




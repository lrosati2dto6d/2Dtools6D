# -*- coding: utf-8 -*-

"""Get Info about ViewTemplates in the model"""

__title__ = "Get ViewTemplates\n Info"
__author__ = "Luca Rosati"

import clr

clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *

clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)

clr.AddReference("RevitServices")

from pyrevit import script

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

def estrai_info_viewtemplates(doc):
    info_templates = []

    # Ottieni tutti i ViewTemplates dal documento corrente
    view_templates = FilteredElementCollector(doc).OfClass(View).ToElements()

    # Ottieni elenco di tutte le categorie di modello disponibili nel documento
    categorie = doc.Settings.Categories
    nomi_vt = []
    for vt in view_templates:
        if not vt.IsTemplate:
            continue

        scala = vt.Scale
        livello_dettaglio = vt.get_Parameter(BuiltInParameter.VIEW_DETAIL_LEVEL).AsValueString()

        # Trova tutte le viste a cui è applicato questo template
        viste_applicate = [v.Name for v in FilteredElementCollector(doc).OfClass(View).ToElements() if v.ViewTemplateId == vt.Id]

        filtri_applicati = []

        # Ottieni gli ID dei filtri applicati al ViewTemplate
        filtro_ids = vt.GetFilters()

        # Per ogni ID, ottieni il filtro corrispondente e le sue informazioni
        for filtro_id in filtro_ids:
            filtro = doc.GetElement(filtro_id)
            if filtro is None:
                continue

            # Ottieni il nome del filtro
            nome_filtro = filtro.Name
            filtri_applicati.append(nome_filtro)

        # Compone informazione in un formato leggibile
        info = "- Scala: {}\n - Livello di Dettaglio: {}\n - Viste Applicate: {}\n - Filtri Applicati: {}".format(scala,livello_dettaglio,(', '.join(viste_applicate)),(', '.join(filtri_applicati)))
        info_templates.append(info)

    return info_templates


view_templates = FilteredElementCollector(doc).OfClass(View).ToElements()


nomi_vt = []
for vt in view_templates:
    if not vt.IsTemplate:
        continue
        
    nomi_vt.append(vt.Name)

r_nomi = range (1, len(nomi_vt))

def mostra_output(info_templates,nomi_def):
    output = script.get_output()
    output.print_md("# VIEW TEMPLATE INFO: {}".format(doc.Title))
    print(' Total Number of View Templates = {}'.format(len(nomi_def)))

    for info,t,n in zip (info_templates,nomi_def,r_nomi):
        output.print_md("## [{}] - {}".format(n,t))
        print("{}".format(info))

info_templates = estrai_info_viewtemplates(doc)

mostra_output(info_templates,nomi_vt)
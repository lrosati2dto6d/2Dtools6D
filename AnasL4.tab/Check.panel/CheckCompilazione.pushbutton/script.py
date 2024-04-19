# -*- coding: utf-8 -*-

""" Esegui diversi passaggi per agevolare la compilazione dei parametri mancanti"""

__title__= 'Compilazione\nAnas L-4'
__author__= 'Luca Rosati'

import clr

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import *

clr.AddReference('RevitNodes')
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(Revit.Elements)

clr.AddReference('RevitServices')
from System.Collections.Generic import *

from pyrevit import forms
from pyrevit import script

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument

rapp = doc.Application


options = ["P1 - Compilazione Di Default","P2 - Verifica Compilare - 999", "P3 - Trasforma ND e 111"]

value_form = forms.ask_for_one_item(
	options,
	default = options[0],
	prompt='Seleziona un passaggio',
	title=',Compilazione L-4')

if value_form == None:
	script.exit()

#-------------------------------------DEFINIZIONI

def ExitScript(check):
	if not check:
		script.exit()

def ConvUnitsFM(number): #Feet to m
	output = number/0.3048
	return output

def ConvUnitsFqMq(number): #Feetq to mq
	output = number/0.092903
	return output

def ParaBuilt (element,builtinparameter):
	parameter = element.get_Parameter(BuiltInParameter.builtinparameter) # type: ignore
	return parameter

def Para(element,paraname):
	parameter = element.LookupParameter(paraname)
	return parameter

def ParaInst(element,paraname):
	if element.LookupParameter(paraname).StorageType == StorageType.Double: # type: ignore
		value = element.LookupParameter(paraname).AsDouble()
	elif element.LookupParameter(paraname).StorageType == StorageType.ElementId: # type: ignore
		value = element.LookupParameter(paraname).AsElementId()
	elif element.LookupParameter(paraname).StorageType == StorageType.String: # type: ignore
		value = element.LookupParameter(paraname).AsString()
	elif element.LookupParameter(paraname).StorageType == StorageType.Integer: # type: ignore
		value = element.LookupParameter(paraname).AsInteger()
	elif element.LookupParameter(paraname).StorageType == None:
		value = "Da Compilare"
	return value

def ParaType(element,paraname,document):
	element_type = document.GetElement(element.GetTypeId())
	if element_type.LookupParameter(paraname).StorageType == StorageType.Double: # type: ignore
		value = element_type.LookupParameter(paraname).AsDouble()
	elif element_type.LookupParameter(paraname).StorageType == StorageType.ElementId: # type: ignore
		value = element_type.LookupParameter(paraname).AsElementId()
	elif element_type.LookupParameter(paraname).StorageType == StorageType.String: # type: ignore
		value = element_type.LookupParameter(paraname).AsString()
	elif element_type.LookupParameter(paraname).StorageType == StorageType.Integer: # type: ignore
		value = element_type.LookupParameter(paraname).AsInteger()
	elif element_type.LookupParameter(paraname).StorageType == None:
		value = "Da Compilare"
	return value

#-------------------------------------FINESTRA OUTPUT

output = script.get_output()

output.resize(1200,800)


#------------------------------------- CHECK_INFORMAZIONI MODELLO
if "FED" in doc.Title:
	file_name = "Nome Modello = {}".format(doc.Title)
else:
	forms.alert('WARNING 00_INFORMAZIONI MODELLO\n\nImpostare correttamente il nome del Modello', exitscript=True)

path_name = "Nome Percorso = {}".format(doc.PathName)

if "2022" in rapp.VersionName:
	rapp_version = "Versione = {}".format(rapp.VersionName)
else:
	forms.alert('WARNING 00_INFORMAZIONI MODELLO\n\nImpostare correttamente la versione di Revit 2022', exitscript=True)


rapp_language = "Lingua = {} :Italy:".format(rapp.Language)
if "Italian" in rapp_language:
	pass
else:
	forms.alert('WARNING 00_INFORMAZIONI MODELLO\n\nImpostare la versione di Revit in lingua Italiana', exitscript=True)

if "Anas_Categorie-IFC Class.txt" in rapp.ExportIFCCategoryTable  and "ACCDocs" in rapp.ExportIFCCategoryTable:
	rapp_ifcclassfile = "File Mappaggio IFC Class = Anas_Categorie-IFC Class.txt"
else:
	forms.alert('WARNING 00_INFORMAZIONI MODELLO\n\nIl file txt del mappaggio delle classi IFC non risulta inserito o non coincide con il percorso condiviso in BIM360', exitscript=True)

if "Anas_ParametriCondivisi_CLUSTER.txt" in rapp.SharedParametersFilename and "ACCDocs" in rapp.SharedParametersFilename:
	rapp_SharedParametersFilename = "File Parametri Condivisi = Anas_ParametriCondivisi_CLUSTER.txt"
else:
	forms.alert('WARNING 00_INFORMAZIONI MODELLO\n\nIl file txt dei parametri condivisi non risulta inserito o non coincide con il percorso condiviso in BIM360', exitscript=True)


info = doc.ProjectInformation

file_info = [file_name,path_name,rapp_version,rapp_language,rapp_ifcclassfile,rapp_SharedParametersFilename]

#------------------------------------- CHECK_EXP.VIEWS

pvp = ParameterValueProvider(ElementId(634609)) # type: ignore
fng = FilterStringEquals() # type: ignore
ruleValue = 'D_ESPORTAZIONI_IFC'
fRule = FilterStringRule(pvp,fng,ruleValue,True) # type: ignore

filter = ElementParameterFilter(fRule) # type: ignore

exp_views_coll = FilteredElementCollector(doc).OfClass(View3D).WherePasses(filter).WhereElementIsNotElementType().ToElements() # type: ignore

exp_views_ele = []

for ev in exp_views_coll:
	if ev.IsTemplate == False:
		exp_views_ele.append(ev)

exp_views = []

for v in exp_views_ele:
	exp_views.append(v.Name)

disc_spec = ["AMB","GET","FED","TRA","IMP","STR"]

exp_view_check = []
exp_view_false = []

for v in exp_views:
	presente = False
	for i in disc_spec:
		if i in v:
			presente = True
			break
	if presente:
		exp_view_check.append("{} --> V".format(v))
	else:
		exp_view_false.append(v)


#WARNING_01-----------------NOME E PRESENZA VISTE ESPORTAZIONE

if len(exp_view_check) == len(exp_views):
	exp_view_result = "WARNING 01_VISTE ESPORTAZIONE IFC --> :white_heavy_check_mark:"
else:
	forms.alert('WARNING 01_VISTE ESPORTAZIONE IFC\n\nLe seguenti viste di esportazione non sono correttemente impostate o rinominate\n\n {}'.format(exp_view_false), exitscript=True)

#-------------------------------------ELEMENTI

for ev in exp_views_ele:
	if "FED" in ev.Name:
		ex_view_fg = ev

#WARNING_02-----------------VISTA ESPORTAZIONE FEDERATO

if ex_view_fg.get_Parameter(BuiltInParameter.VIEW_PHASE).AsValueString()=="Nuova costruzione" and ex_view_fg.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER).AsValueString()=="Mostra completo":
	exp_viewed_result = "WARNING 02_VISTA EXP FEDERATO --> :white_heavy_check_mark:"
else:
	forms.alert('WARNING 02_VISTA EXP FEDERATO\n\nLa vista di esportazione FED non ha le fasi impostate correttamente\n\n Correggere la vista con le seguenti impostazioni\n\nFiltro delle fasi = Mostra completo\nFase = Nuova costruzione', exitscript=True)

doc_el = FilteredElementCollector(doc,ex_view_fg.Id).WhereElementIsNotElementType().ToElements() # type: ignore


cat = []
cnamelist = []
catname = []

clean_el = []

for el in doc_el:
	try:
		if el.Category.CategoryType == CategoryType.Model and "dwg"  not in el.Category.Name and el.Category.SubCategories.Size > 0 and el.Category.CanAddSubcategory: # type: ignore
			clean_el.append(el)
	except:
		pass

if len(clean_el) == 0:
	forms.alert('WARNING 02_VISTA EXP FEDERATO\n\nLa vista di esportazione FED non contiene nessun elemento\n\n Correggere  le impostazioni della vista', exitscript=True)

for el in clean_el:
	catname.append(el.Category.Name)
	cat.append(el.Category)
	cnamelist.append(el)

c_uni = []

for ci in catname:
	if ci not in c_uni:
		c_uni.append(ci)

n_el= "Numero Totale Elementi = {}".format(len(cnamelist))

cat_true_name = ("Apparecchi elettrici", "Appoggi", "Armatura strutturale", "Attrezzatura elettrica", "Attrezzature speciali", "Tubazioni flessibili", "Sistema di tubazioni", "Collegamenti strutturali", "Fondazioni strutturali", "Modelli generici", "Muri", "Pavimenti", "Pilastri strutturali", "Piloni", "Ringhiere", "Telaio strutturale", "Tetti")
cat_false = []

for c in c_uni:
	if c in cat_true_name:
		cat_result = "WARNING 03_CATEGORIE MODELLO DATI --> :white_heavy_check_mark:"
	else:
		cat_false.append(c)

#WARNING_03-----------------CATEGORIA

if len(cat_false) != 0:
	forms.alert('WARNING 03_CATEGORIE\n\nLe seguenti categorie non sono presenti nel modello dati,\n\n{}\n\nEliminare gli elementi prima di procedere'.format(cat_false), exitscript=True)


#-------------------------------------ASSOCIAZIONE FASE

numfe = 0
assfase_errata = []

for el in clean_el:
	type_el = doc.GetElement(el.GetTypeId())
	category_el = el.Category.Name
	opera_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsValueString()
	parteopera_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString()
	elemento_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsValueString()
	el_id = el.Id
	try:
		symbol = el.Symbol
		type_el_name = type_el.FamilyName
	except:
		type_el_name = el.Name

	if opera_el in ['IM','MO'] and el.get_Parameter(BuiltInParameter.PHASE_CREATED).AsValueString() != "Nuova costruzione":
		numfe += 1
		assfase_errata.append(":heavy_multiplication_x: {} - {} - {} - {} - {} - {} - {}".format(num,category_el,type_el_name,opera_el,parteopera_el,elemento_el,output.linkify(el_id)))

	elif opera_el in ['PV','MA','SI'] and el.get_Parameter(BuiltInParameter.PHASE_CREATED).AsValueString() != "Esistente":
		numfe += 1
		assfase_errata.append(":heavy_multiplication_x: {} - {} - {} - {} - {} - {} - {}".format(numfe,category_el,type_el_name,opera_el,parteopera_el,elemento_el,output.linkify(el_id)))


#VERIFICA_01-----------------ASSOCIAZIONE FASE

if len(assfase_errata) != 0:
	output.print_md(	'# :red_circle: VERIFICA 01_ASSOCIAZIONE ELEMENTI-FASE')
	output.print_md(	'##I seguenti Elementi hanno una Associazione della Fase di Creazione Errata')
	for i in assfase_errata:
		output.print_md(	'###{}'.format(i))
else:
	result_afase = "VERIFICA 01_ASSOCIAZIONE ELEMENTI-FASE --> :white_heavy_check_mark:"

if len(assfase_errata) != 0:
	script.exit()

#-------------------------------------ASSOCIAZIONE WORKSET

numfe = 0
worksetfase_errata = []

for el in clean_el:
	type_el = doc.GetElement(el.GetTypeId())
	category_el = el.Category.Name
	opera_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsValueString()
	parteopera_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString()
	elemento_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsValueString()
	el_id = el.Id
	try:
		symbol = el.Symbol
		type_el_name = type_el.FamilyName
	except:
		type_el_name = el.Name

	if opera_el in ['IM','MO'] and el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString() != "08_IMP":
		numfe += 1
		worksetfase_errata.append(":heavy_multiplication_x: {} - {} - {} - {} - {} - {} - {}".format(numfe,category_el,type_el_name,opera_el,parteopera_el,elemento_el,output.linkify(el_id)))

	elif elemento_el in ['BAN','PAN'] and el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString() != "07_AMB":
		numfe += 1
		worksetfase_errata.append(":heavy_multiplication_x: {} - {} - {} - {} - {} - {} - {}".format(numfe,category_el,type_el_name,opera_el,parteopera_el,elemento_el,output.linkify(el_id)))

	elif elemento_el in ['CEN','MON','MPL','PAL','PLI','POZ','PZF','RAN'] and el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString() != "06_GET":
		numfe += 1
		worksetfase_errata.append(":heavy_multiplication_x: {} - {} - {} - {} - {} - {} - {}".format(numfe,category_el,type_el_name,opera_el,parteopera_el,elemento_el,output.linkify(el_id)))

	elif elemento_el in ['BIN','LMC','MUS','NJE','PPZ','TTA','UNI'] and el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString() != "05_TRA":
		numfe += 1
		worksetfase_errata.append(":heavy_multiplication_x: {} - {} - {} - {} - {} - {} - {}".format(numfe,category_el,type_el_name,opera_el,parteopera_el,elemento_el,output.linkify(el_id)))

	elif elemento_el in ['AAP','AMS','BAG','BPT','CAP','CAS','CNT','COR','CUN','DIA','GIU','ISA','LOR','MAN','MDA','MFR','OPO','PAR','PEN','PUL','PUN','RIS','SAR','SBL','SEL','SGE','SOL','SSB','STL','TAN','TIM','TRA','TRV','VEL'] and el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString() != "04_STR":
		numfe += 1
		worksetfase_errata.append(":heavy_multiplication_x: {} - {} - {} - {} - {} - {} - {}".format(numfe,category_el,type_el_name,opera_el,parteopera_el,elemento_el,output.linkify(el_id)))



#VERIFICA_02-----------------ASSOCIAZIONE WORKSET

if len(worksetfase_errata) != 0:
	output.print_md(	'# :red_circle: VERIFICA 02_ASSOCIAZIONE ELEMENTI-WORKSET')
	output.print_md(	'##I seguenti Elementi hanno una Associazione del Workset Errata')
	for i in worksetfase_errata:
		output.print_md(	'###{}'.format(i))
else:
	result_workset = "VERIFICA 02_ASSOCIAZIONE ELEMENTI-WORKSET --> :white_heavy_check_mark:"

if len(worksetfase_errata) != 0:
	script.exit()


#-------------------------------------CLASSIFICAZIONE

num = 0
class_CVPV = []
class_IM = []
class_MA = []
class_MO = []
class_SI = []

class_errata =[]

for el in clean_el:
	type_el = doc.GetElement(el.GetTypeId())
	category_el = el.Category.Name
	opera_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsValueString()
	parteopera_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString()
	elemento_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsValueString()
	el_id = el.Id
	try:
		symbol = el.Symbol
		type_el_name = type_el.FamilyName
	except:
		type_el_name = el.Name

	if opera_el == 'PV' and category_el in ['Telaio strutturale', 'Modelli generici', 'Pilastri strutturali', 'Tetti', 'Appoggi', 'Fondazioni strutturali', 'Collegamenti strutturali', 'Piloni', 'Armatura strutturale', 'Pavimenti', 'Muri'] and parteopera_el in ['AC', 'AN', 'FO', 'FP', 'FS', 'GI', 'IC','IA', 'PD', 'PI', 'SP'] and elemento_el in ['AAP', 'AMS', 'BAG', 'BIN', 'CAP', 'CAS', 'CEN', 'CNT', 'COR', 'CUN', 'DIA', 'GIU', 'ISA', 'LOR', 'MAN', 'MDA', 'MFR', 'MPL', 'MUS', 'OPO', 'PAL', 'PAR', 'PEN', 'POZ', 'PPZ', 'PUL', 'PUN', 'PZF', 'RIS', 'SAR', 'SBL', 'SEL', 'SGE', 'SOL', 'SSG', 'STL', 'TAN', 'TIM', 'TRA', 'TRV', 'VEL']:
		class_CVPV = True

	elif opera_el == 'IM' and category_el in ['Apparecchi elettrici', 'Attrezzatura elettrica', 'Attrezzature speciali', 'Tubazioni flessibili', 'Collegamenti strutturali', 'Fondazioni strutturali']and parteopera_el =='IE'and elemento_el in ['ALI', 'BLI', 'CAE', 'CAV', 'CEE', 'CNP', 'COL', 'COM', 'DIE', 'IMT', 'INM', 'INS', 'MUL', 'PLI', 'PZE', 'QEB', 'QEM', 'REL', 'REP', 'RIF', 'SCS', 'SDE', 'TRS']:
		class_IM = True

	elif opera_el == 'MA' and category_el in ['Fondazioni strutturali', 'Modelli generici', 'Ringhiere', 'Telaio strutturale'] and parteopera_el =='BA' and elemento_el in ['BAN', 'BPT', 'COR', 'MON']:
		class_MA = True

	elif opera_el == 'MO' and category_el in ['Apparecchi elettrici', 'Modelli generici'] and parteopera_el in ['ES', 'MN', 'TR', 'BE', 'CP', 'EA'] and elemento_el in ['ACC', 'IFS', 'SCA', 'SEM', 'TCM', 'TEC', 'VCM','BRE', 'CLI', 'IDO', 'SUM', 'TIG', 'TSS']:
		class_MO = True

	elif opera_el == 'SI' and category_el in ['Ringhiere', 'Modelli generici', 'Collegamenti strutturali'] and parteopera_el in ['BC','BS'] and elemento_el in ['LMC', 'NJE', 'PAN', 'RAN', 'TTA', 'UNI']:
		class_SI = True

	else:
		num += 1
		class_errata.append(":heavy_multiplication_x: {} - {} - {} - {} - {} - {} - {}".format(num,category_el,type_el_name,opera_el,parteopera_el,elemento_el,output.linkify(el_id)))


#VERIFICA_03-----------------CLASSIFICAZIONE

if len(class_errata) != 0:
	output.print_md(	'# :red_circle: VERIFICA 03_CLASSIFICAZIONE ELEMENTI')
	output.print_md(	'## I seguenti Elementi hanno una Classificazione errata secondo il Modello Dati')
	for i in class_errata:
		output.print_md(	'###{}'.format(i))

else:
	result_class = "VERIFICA 03_CLASSIFICAZIONE ELEMENTI --> :white_heavy_check_mark:"

if len(class_errata) != 0:
	script.exit()


#-------------------------------------NOMENCLATURA

num_n = 0
num_h = 0
nomen_errata = []
host_errato = []

for el in clean_el:
	type_el = doc.GetElement(el.GetTypeId())
	category_el = el.Category.Name
	opera_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsValueString()
	parteopera_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString()
	elemento_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsValueString()
	el_id = el.Id
	try:
		symbol = el.Symbol
		type_el_name = type_el.FamilyName
	except:
		type_el_name = el.Name

	if opera_el in type_el_name and parteopera_el in type_el_name and elemento_el in type_el_name:
		Nomen = True
	else:
		num_n += 1
		nomen_errata.append("{} - {} - {} - {} - {} - {} - {}".format(num_n,category_el,type_el_name,opera_el,parteopera_el,elemento_el,output.linkify(el_id)))

#VERIFICA_04-----------------NOMENCLATURA

if len(nomen_errata) != 0:
	output.print_md(	'# :red_circle: VERIFICA 04_NOMENCLATURA ELEMENTI')
	output.print_md(	'## I seguenti Elementi hanno una Nomenclatura errata secondo il Modello Dati')
	for i in nomen_errata:
		output.print_md(	'###{}'.format(i))

else:
	result_nomen = "VERIFICA 04_NOMENCLATURA ELEMENTI --> :white_heavy_check_mark:"

if len(nomen_errata) != 0:
	script.exit()

#-------------------------------------PASSAGGIO 1 - COMPILA NECESSARI

param_inst_compilare = []
inst_compilare = []

param_typ_compilare = []
typ_compilare = []

param_inst_999 = []
inst_999 = []

parastrano = doc.GetElement(ElementId(637674))

#-------------------------------------PASSAGGIO 2 - CHECK 

para_inst_check = []
para_typ_check = set()
inst_check = []
typ_check= []

#----- ELIMINA ECCESSO

para_eccesso = []

para_num_eccesso = []

#-------------------------------------PASSAGGIO 3 - TRSFORMA NECESSARI

para_ND_trasforma = []

para_111_trasforma = []

info = doc.ProjectInformation

#-----IDP_Identificativo Progetto(txt_inst)

infoparameters = ["IDP_AGR","IDP_CIG","IDP_Codice Intervento","IDP_Codice PPM","IDP_CUP","IDP_Denominazione","IDP_Fase progettuale","IDP_Nome opera","IDP_Struttura territoriale","RES_BIM Manager Anas","RES_Esecutore modellazione","RES_Impresa Esecutrice","RES_Progettista","RES_Responsabile Unico del Procedimento (RUP)","SIT_Codice Strada","SIT_Comune","SIT_Regione","SIT_Sistema di Coordinate"]

for ip in infoparameters:
	if Para(info,ip).HasValue == False or ParaInst(info,ip) == "" or ParaInst(info,ip) == None:
		param_inst_compilare.append(Para(info,ip))
		inst_compilare.append(info)
	elif ParaInst(info,ip) == "•••COMPILARE•••":
		para_inst_check.append("{} --> :heavy_multiplication_x:".format(Para(info,ip).Definition.Name))
	elif ParaInst(info,ip) == "ND":
		para_ND_trasforma.append(Para(info,ip))
	else:
		pass


for el in clean_el:
	type_el = doc.GetElement(el.GetTypeId())
	category_el = el.Category.Name
	opera_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsValueString()
	parteopera_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString()
	elemento_el = type_el.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsValueString()
	el_id = el.Id
	try:
		symbol = el.Symbol
		type_el_name = type_el.FamilyName
	except:
		type_el_name = el.Name



#-----ANA_Esecutore (txt_inst)

	if opera_el in ["CV","IM","MA","MO","PV","SI"] and parteopera_el in ["AN","AC","FP","FO","FS","GI","IC","IA","PD","PI","SP","IE","BA","ES","BS","BC"] and elemento_el in ["STL","SEL","TRV","CAP","OPO","PEN","SAR","TIM","DIA","MPL","PAL","PZF","CEN","POZ","GIU","AAP","BIN","CAS","COR","CUN","MUS","PPZ","SGE","SOL","SSB","TRA","VEL","CNT","LOR","PUN","SBL","TAN","AMS","BAG","ISA","PUL","RIS","MAN","MDA","MFR","PAR","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","CEE","DIE","IMT","INS","INM","MUL","PLI","PZE","QEB","QEM","REL","REP","RIF","SCS","SDE","TRS","BAN","BPT","MON","TCM","LMC","PAN","RAN","TTA","UNI","NJE"]:
		if Para(el,"ANA_Esecutore").HasValue == False or ParaInst(el,"ANA_Esecutore") == "" or ParaInst(el,"ANA_Esecutore") == None:
			param_inst_compilare.append(Para(el,"ANA_Esecutore"))
			inst_compilare.append(el)
		elif ParaInst(el,"ANA_Esecutore") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"ANA_Esecutore").Definition.Name))
		elif ParaInst(el,"ANA_Esecutore") == "ND":
			para_ND_trasforma.append(Para(el,"ANA_Esecutore"))
	else:
		try:
			para_eccesso.append(Para(el,"ANA_Esecutore"))
		except:
			pass


#-----ANA_Progettista (txt_inst)

	if opera_el in ["CV","IM","MA","PV","SI"] and parteopera_el in ["AN","AC","FP","FO","FS","GI","IC","IA","PD","PI","SP","IE","BA","BS","BC"] and elemento_el in ["STL","SEL","TRV","CAP","OPO","PEN","SAR","TIM","DIA","MPL","PAL","PZF","CEN","POZ","GIU","AAP","BIN","CAS","COR","CUN","MUS","PPZ","SGE","SOL","SSB","TRA","VEL","CNT","LOR","PUN","SBL","TAN","AMS","BAG","ISA","PUL","RIS","MAN","MDA","MFR","PAR","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","CEE","DIE","IMT","INS","INM","MUL","PLI","PZE","QEB","QEM","REL","REP","RIF","SCS","SDE","TRS","BAN","BPT","MON","LMC","PAN","RAN","TTA","UNI","NJE"]:
		if Para(el,"ANA_Progettista").HasValue == False or ParaInst(el,"ANA_Progettista") == "" or ParaInst(el,"ANA_Progettista") == None:
			param_inst_compilare.append(Para(el,"ANA_Progettista"))
			inst_compilare.append(el)
		elif ParaInst(el,"ANA_Progettista") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"ANA_Progettista").Definition.Name))
		elif ParaInst(el,"ANA_Progettista") == "ND":
			para_ND_trasforma.append(Para(el,"ANA_Progettista"))
	else:
		try:
			para_eccesso.append(Para(el,"ANA_Progettista"))
		except:
			pass


#-------------------------------------GEOMETRICO

#-----GEO_Identificativo concio (txt_inst)

	if opera_el in ["CV","MA","PV"] and parteopera_el in ["AN","AC","FP","FO","FS","IC","IA","PD","PI","SP","BA"] and elemento_el in ["TRV","MPL","PAL","PZF","CEN","COR","SOL","TRA"]:
		if Para(el,"GEO_Identificativo concio").HasValue == False or ParaInst(el,"GEO_Identificativo concio") == "" or ParaInst(el,"GEO_Identificativo concio") == None:
			param_inst_compilare.append(Para(el,"GEO_Identificativo concio"))
			inst_compilare.append(el)
		elif ParaInst(el,"GEO_Identificativo concio") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Identificativo concio").Definition.Name))
		elif ParaInst(el,"GEO_Identificativo concio") == "ND":
			para_ND_trasforma.append(Para(el,"GEO_Identificativo concio"))
	else:
		try:
			para_eccesso.append(Para(el,"GEO_Identificativo concio"))
		except:
			pass

#-----GEO_Altezza (len_inst)
	if opera_el in ["CV","IM","MA","MO","PV","SI"] and parteopera_el in ["AN","AC","FP","FO","FS","GI","IC","IA","PD","PI","SP","IE","BA","BE","CP","MN","BS","BC"] and elemento_el in ["SEL","CAP","OPO","PEN","TIM","PZF","GIU","AAP","CAS","COR","CUN","SGE","TRA","VEL","AMS","BAG","ISA","PUL","MAN","MDA","MFR","PAR","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","CEE","DIE","IMT","INS","INM","MUL","PLI","QEB","QEM","REL","REP","RIF","SCS","SDE","TRS","BAN","BPT","BRE","CLI","ACC","SCA","VCM","LMC","PAN","RAN","TTA","NJE"]:
		if Para(el,"GEO_Altezza").HasValue == False or ParaInst(el,"GEO_Altezza") == None:
			param_inst_999.append(Para(el,"GEO_Altezza"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Altezza") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Altezza").Definition.Name))
		elif ParaInst(el,"GEO_Altezza") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Altezza"))
	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Altezza"))
		except:
			pass

#-----GEO_Altezza interna (len_inst)
	if opera_el in ["CV","IM","PV"] and parteopera_el in ["IC","IA","IE"] and elemento_el in ["CAS","SGE","TRA","PZE"]:
		if Para(el,"GEO_Altezza interna").HasValue == False or ParaInst(el,"GEO_Altezza interna") == None:
			param_inst_999.append(Para(el,"GEO_Altezza interna"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Altezza interna") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Altezza interna").Definition.Name))
		elif ParaInst(el,"GEO_Altezza interna") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Altezza interna"))
	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Altezza interna"))
		except:
			pass

#-----GEO_Altezza profilo (len_inst)
	if opera_el in ["CV","MA","PV"] and parteopera_el in ["AN","AC","FO","IC","IA","PD","PI","SP","BA"] and elemento_el in ["TRV","CEN","CNT","LOR","PUN","SBL","TAN","RIS","MON"]:
		if Para(el,"GEO_Altezza profilo").HasValue == False or ParaInst(el,"GEO_Altezza profilo") == None:
			param_inst_999.append(Para(el,"GEO_Altezza profilo"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Altezza profilo") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Altezza profilo").Definition.Name))
		elif ParaInst(el,"GEO_Altezza profilo") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Altezza profilo"))
	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Altezza profilo"))
		except:
			pass

#-----GEO_Altezza sezione (len_inst)
#-----GEO_Larghezza sezione (len_inst)
	if opera_el in ["CV","PV"] and parteopera_el in ["AC","FP","FO","FS"] and elemento_el in ["SAR","PZF"]:
		if Para(el,"GEO_Altezza sezione").HasValue == False or ParaInst(el,"GEO_Altezza sezione") == None:
			param_inst_999.append(Para(el,"GEO_Altezza sezione"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Altezza sezione") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Altezza sezione").Definition.Name))
		elif ParaInst(el,"GEO_Altezza sezione") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Altezza sezione"))

		if Para(el,"GEO_Larghezza sezione").HasValue == False or ParaInst(el,"GEO_Larghezza sezione") == None:
			param_inst_999.append(Para(el,"GEO_Larghezza sezione"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Larghezza sezione") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Larghezza sezione").Definition.Name))
		elif ParaInst(el,"GEO_Larghezza sezione") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Larghezza sezione"))
	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Altezza sezione"))
			para_num_eccesso.append(Para(el,"GEO_Larghezza sezione"))
		except:
			pass

#-----GEO_Area (len_inst)
	if opera_el in ["CV","IM","PV","SI"] and parteopera_el in ["FP","FO","FS","IC","IA","SP","IE","BS"] and elemento_el in ["PZF","CEN","POZ","SOL","SSB","CEE","PLI","QEB","QEM","RAN"]:
		if Para(el,"GEO_Area").HasValue == False or ParaInst(el,"GEO_Area") == None:
			param_inst_999.append(Para(el,"GEO_Area"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Area") == ConvUnitsFqMq(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Area").Definition.Name))
		elif Para(el,"GEO_Area").AsValueString() == "111 m²":
			para_111_trasforma.append(Para(el,"GEO_Area"))
	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Area"))
		except:
			pass

#-----GEO_Diametro (len__inst)
	if opera_el in ["CV","IM","MO","PV"] and parteopera_el in ["AN","AC","FP","FO","IC","IA","PD","PI","SP","IE","EA","ES","TR"] and elemento_el in ["STL","PEN","DIA","MPL","PAL","CEN","POZ","AAP","AMS","ISA","BLI","CNP","CAE","CAV","DIE","IMT","PZE","TSS","IFS","TEC","SEM","SUM","TIG"]:
		if Para(el,"GEO_Diametro").HasValue == False or ParaInst(el,"GEO_Diametro") == None:
			param_inst_999.append(Para(el,"GEO_Diametro"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Diametro") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Diametro").Definition.Name))
		elif ParaInst(el,"GEO_Diametro") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Diametro"))
	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Diametro"))
		except:
			pass

#-----GEO_Diametro esterno (len__inst)
#-----GEO_Diametro interno (len__inst)
	if opera_el in ["CV","PV"] and parteopera_el in ["AN","PD","PI"] and elemento_el in ["SEL"]:
		if Para(el,"GEO_Diametro esterno").HasValue == False or ParaInst(el,"GEO_Diametro esterno") == None:
			param_inst_999.append(Para(el,"GEO_Diametro esterno"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Diametro esterno") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Diametro esterno").Definition.Name))
		elif ParaInst(el,"GEO_Diametro esterno") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Diametro esterno"))

		if Para(el,"GEO_Diametro interno").HasValue == False or ParaInst(el,"GEO_Diametro interno") == None:
			param_inst_999.append(Para(el,"GEO_Diametro interno"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Diametro interno") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Diametro esterno").Definition.Name))
		elif ParaInst(el,"GEO_Diametro interno") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Diametro interno"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Diametro esterno"))
			para_num_eccesso.append(Para(el,"GEO_Diametro interno"))
		except:
			pass

#-----GEO_Diametro tubo (len__inst)
	if opera_el in ["MA"] and parteopera_el in ["BA"] and elemento_el in ["MON"]:
		if Para(el,"GEO_Diametro tubo").HasValue == False or ParaInst(el,"GEO_Diametro tubo") == None:
			param_inst_999.append(Para(el,"GEO_Diametro tubo"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Diametro tubo") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Diametro tubo").Definition.Name))
		elif ParaInst(el,"GEO_Diametro tubo") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Diametro tubo"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Diametro tubo"))
		except:
			pass

#-----GEO_Larghezza (len__inst)
	if opera_el in ["CV","IM","MA","MO","PV"] and parteopera_el in ["AN","AC","FP","FO","FS","GI","IC","IA","PD","PI","SP","IE","BA","BE","CP","MN","TR"] and elemento_el in ["SEL","CAP","OPO","PEN","TIM","DIA","MPL","PAL","PZF","GIU","AAP","CAS","COR","CUN","SGE","SOL","SSB","TRA","VEL","AMS","BAG","ISA","PUL","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","CEE","DIE","IMT","INS","INM","MUL","PLI","QEB","QEM","REL","REP","RIF","SCS","SDE","TRS","BAN","BPT","BRE","CLI","ACC","SCA","VCM","IDO"]:
		if Para(el,"GEO_Larghezza").HasValue == False or ParaInst(el,"GEO_Larghezza") == None:
			param_inst_999.append(Para(el,"GEO_Larghezza"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Larghezza") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Larghezza").Definition.Name))
		elif ParaInst(el,"GEO_Larghezza") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Larghezza"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Larghezza"))
		except:
			pass

#-----GEO_Larghezza interna (len__inst)
	if opera_el in ["CV","IM","PV"] and parteopera_el in ["IC","IA","IE"] and elemento_el in ["CAS","SGE","TRA","PZE"]:
		if Para(el,"GEO_Larghezza interna").HasValue == False or ParaInst(el,"GEO_Larghezza interna") == None:
			param_inst_999.append(Para(el,"GEO_Larghezza interna"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Larghezza interna") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Larghezza interna").Definition.Name))
		elif ParaInst(el,"GEO_Larghezza interna") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Larghezza interna"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Larghezza interna"))
		except:
			pass

#-----GEO_Larghezza profilo (len__inst)
#-----GEO_Lunghezza profilo (len__inst)
	if opera_el in ["CV","MA","PV","SI"] and parteopera_el in ["AN","AC","FO","IC","IA","PD","PI","SP","BA","BS","BC"] and elemento_el in ["TRV","CEN","CNT","LOR","PUN","SBL","TAN","RIS","MON","UNI"]:
		if Para(el,"GEO_Larghezza profilo").HasValue == False or ParaInst(el,"GEO_Larghezza profilo") == None:
			param_inst_999.append(Para(el,"GEO_Larghezza profilo"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Larghezza profilo") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Larghezza profilo").Definition.Name))
		elif ParaInst(el,"GEO_Larghezza profilo") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Larghezza profilo"))

			
		if Para(el,"GEO_Lunghezza profilo").HasValue == False or ParaInst(el,"GEO_Lunghezza profilo") == None:
			param_inst_999.append(Para(el,"GEO_Lunghezza profilo"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Lunghezza profilo") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Lunghezza profilo").Definition.Name))
		elif ParaInst(el,"GEO_Lunghezza profilo") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Lunghezza profilo"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Larghezza profilo"))
			para_num_eccesso.append(Para(el,"GEO_Lunghezza profilo"))
		except:
			pass

#-----GEO_Luce (len__inst)
	if opera_el in ["CV","PV"] and parteopera_el in ["FO"] and elemento_el in ["CEN"]:
		if Para(el,"GEO_Luce").HasValue == False or ParaInst(el,"GEO_Luce") == None:
			param_inst_999.append(Para(el,"GEO_Luce"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Luce") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Luce").Definition.Name))
		elif ParaInst(el,"GEO_Luce") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Luce"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Luce"))
		except:
			pass

#-----GEO_Lunghezza (len__inst)
	if opera_el in ["CV","IM","MA","MO","PV","SI"] and parteopera_el in ["AN","AC","FP","FO","FS","GI","IC","IA","PD","PI","SP","IE","BA","BE","CP","EA","ES","MN","TR","BS","BC"] and elemento_el in ["STL","SEL","CAP","OPO","PEN","SAR","TIM","DIA","MPL","PAL","PZF","CEN","POZ","GIU","AAP","CAS","COR","CUN","SGE","SOL","SSB","TRA","VEL","AMS","BAG","ISA","PUL","MAN","MDA","MFR","PAR","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","DIE","IMT","INS","INM","PLI","REL","REP","RIF","SCS","SDE","TRS","BAN","BPT","BRE","CLI","TSS","TEC","ACC","SCA","VCM","IDO","SEM","SUM","TIG","LMC","PAN","RAN","TTA","NJE"]:
		if Para(el,"GEO_Lunghezza").HasValue == False or ParaInst(el,"GEO_Lunghezza") == None:
			param_inst_999.append(Para(el,"GEO_Lunghezza"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Lunghezza") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Lunghezza").Definition.Name))
		elif ParaInst(el,"GEO_Lunghezza") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Lunghezza"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Lunghezza"))
		except:
			pass

#-----GEO_Ingombro trasversale (len__inst)
	if opera_el in ["SI"] and parteopera_el in ["BS","BC"] and elemento_el in ["LMC","TTA","NJE"]:
		if Para(el,"GEO_Ingombro trasversale").HasValue == False or ParaInst(el,"GEO_Ingombro trasversale") == None:
			param_inst_999.append(Para(el,"GEO_Ingombro trasversale"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Ingombro trasversale") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Ingombro trasversale").Definition.Name))
		elif ParaInst(el,"GEO_Ingombro trasversale") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Ingombro trasversale"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Ingombro trasversale"))
		except:
			pass


#-----GEO_Lunghezza interna (len__inst)
	if opera_el in ["IM"] and parteopera_el in ["IE"] and elemento_el in ["PZE"]:
		if Para(el,"GEO_Lunghezza interna").HasValue == False or ParaInst(el,"GEO_Lunghezza interna") == None:
			param_inst_999.append(Para(el,"GEO_Lunghezza interna"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Lunghezza interna") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Lunghezza interna").Definition.Name))
		elif ParaInst(el,"GEO_Lunghezza interna") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Lunghezza interna"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Lunghezza interna"))
		except:
			pass

#-----GEO_Lunghezza totale della catena (len__inst)
	if opera_el in ["MO"] and parteopera_el in ["ES"] and elemento_el in ["IFS"]:
		if Para(el,"GEO_Lunghezza totale della catena").HasValue == False or ParaInst(el,"GEO_Lunghezza totale della catena") == None:
			param_inst_999.append(Para(el,"GEO_Lunghezza totale della catena"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Lunghezza totale della catena") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Lunghezza totale della catena").Definition.Name))
		elif ParaInst(el,"GEO_Lunghezza totale della catena") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Lunghezza totale della catena"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Lunghezza totale della catena"))
		except:
			pass

#-----GEO_Pendenza (slope__inst)
	if opera_el in ["CV","MA","PV"] and parteopera_el in ["IC","IA","SP","BA"] and elemento_el in ["COR","CUN","SOL","SSB","VEL","BPT"]:
		if Para(el,"GEO_Pendenza").HasValue == False or ParaInst(el,"GEO_Pendenza") == None:
			param_inst_999.append(Para(el,"GEO_Pendenza"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Pendenza") == 57.289961630754:
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Pendenza").Definition.Name))
		elif Para(el,"GEO_Pendenza").AsValueString() == "11.000°":
			para_111_trasforma.append(Para(el,"GEO_Pendenza"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Pendenza"))
		except:
			pass

#-----GEO_Pendenza sponde (slope__inst)
	if opera_el in ["CV","MA","PV"] and parteopera_el in ["IC","IA","BA"] and elemento_el in ["COR","CUN","VEL","BPT"]:
		if Para(el,"GEO_Pendenza sponde").HasValue == False or ParaInst(el,"GEO_Pendenza sponde") == None:
			param_inst_999.append(Para(el,"GEO_Pendenza sponde"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Pendenza sponde") == 57.289961630754:
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Pendenza sponde").Definition.Name))
		elif Para(el,"GEO_Pendenza sponde").AsValueString() == "11.000°":
			para_111_trasforma.append(Para(el,"GEO_Pendenza sponde"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Pendenza sponde"))
		except:
			pass

#-----GEO_Peso (weight__inst)
	if opera_el in ["CV","IM","MA","PV","SI"] and parteopera_el in ["AN","AC","FP","FO","FS","GI","IC","IA","PD","PI","SP","IE","BA","BS","BC"] and elemento_el in ["STL","SEL","TRV","PEN","SAR","PZF","CEN","POZ","GIU","AAP","CAS","COR","CUN","SGE","SOL","TRA","VEL","CNT","LOR","PUN","SBL","TAN","AMS","BAG","ISA","PUL","RIS","MAN","MDA","MFR","PAR","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","DIE","IMT","INS","INM","MUL","PLI","PZE","REL","REP","RIF","SCS","SDE","TRS","BAN","BPT","MON","LMC","PAN","RAN","TTA","NJE"]:
		if Para(el,"GEO_Peso").HasValue == False or ParaInst(el,"GEO_Peso") == None:
			param_inst_999.append(Para(el,"GEO_Peso"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Peso") == 3277559.05511811:
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Peso").Definition.Name))
		elif Para(el,"GEO_Peso").AsValueString() == "111.00 kN":
			para_111_trasforma.append(Para(el,"GEO_Peso"))
	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Peso"))
		except:
			pass

#-----GEO_Quota altimetrica inferiore (len__inst)
	if opera_el in ["CV","IM","MA","PV"] and parteopera_el in ["AN","AC","FP","FO","FS","IC","IA","PD","PI","SP","IE","BA"] and elemento_el in ["SEL","SAR","DIA","MPL","PAL","PZF","POZ","AAP","COR","SOL","AMS","BAG","ISA","PUL","MAN","MDA","MFR","PAR","PLI"]:
		if Para(el,"GEO_Quota altimetrica inferiore").HasValue == False or ParaInst(el,"GEO_Quota altimetrica inferiore") == None:
			param_inst_999.append(Para(el,"GEO_Quota altimetrica inferiore"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Quota altimetrica inferiore") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Quota altimetrica inferiore").Definition.Name))
		elif ParaInst(el,"GEO_Quota altimetrica inferiore") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Quota altimetrica inferiore"))
	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Quota altimetrica inferiore"))
		except:
			pass

#-----GEO_Quota altimetrica superiore (len__inst)
	if opera_el in ["CV","PV"] and parteopera_el in ["AN","AC","FP","FO","IC","IA","PD","PI","SP"] and elemento_el in ["SEL","SAR","DIA","MPL","PAL","POZ","AAP","SOL","AMS","BAG","ISA","PUL","MAN","MDA","MFR","PAR"]:
		if Para(el,"GEO_Quota altimetrica superiore").HasValue == False or ParaInst(el,"GEO_Quota altimetrica superiore") == None:
			param_inst_999.append(Para(el,"GEO_Quota altimetrica superiore"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Quota altimetrica superiore") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Quota altimetrica superiore").Definition.Name))
		elif ParaInst(el,"GEO_Quota altimetrica superiore") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Quota altimetrica superiore"))
	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Quota altimetrica superiore"))
		except:
			pass

#-----GEO_Quota fine tratto (len__inst)
#-----GEO_Quota inizio tratto (len__inst)
	if opera_el in ["CV","PV"] and parteopera_el in ["IC","IA","BA"] and elemento_el in ["COR","CUN","VEL","BPT"]:
		if Para(el,"GEO_Quota fine tratto").HasValue == False or ParaInst(el,"GEO_Quota fine tratto") == None:
			param_inst_999.append(Para(el,"GEO_Quota fine tratto"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Quota fine tratto") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Quota fine tratto").Definition.Name))
		elif ParaInst(el,"GEO_Quota fine tratto") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Quota fine tratto"))

		if Para(el,"GEO_Quota inizio tratto").HasValue == False or ParaInst(el,"GEO_Quota inizio tratto") == None:
			param_inst_999.append(Para(el,"GEO_Quota inizio tratto"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Quota inizio tratto") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Quota inizio tratto").Definition.Name))
		elif ParaInst(el,"GEO_Quota inizio tratto") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Quota inizio tratto"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Quota fine tratto"))
			para_num_eccesso.append(Para(el,"GEO_Quota inizio tratto"))
		except:
			pass

#-----GEO_Quota sensore (len__inst)
	if opera_el in ["MO"] and parteopera_el in ["BE","CP","EA","ES","MN","TR"] and elemento_el in ["BRE","CLI","TSS","IFS","TEC","TCM","ACC","SCA","VCM","IDO","SEM","SUM","TIG"]:
		if Para(el,"GEO_Quota sensore").HasValue == False or ParaInst(el,"GEO_Quota sensore") == 0 or ParaInst(el,"GEO_Quota sensore") == None:
			param_inst_999.append(Para(el,"GEO_Quota sensore"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Quota sensore") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Quota sensore").Definition.Name))
	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Quota sensore"))
		except:
			pass

#-----GEO_Raggio (len__inst)
	if opera_el in ["CV","MA","PV"] and parteopera_el in ["AC","IC","IA","BA"] and elemento_el in ["SAR","COR","CUN","VEL","BPT"]:
		if Para(el,"GEO_Raggio").HasValue == False or ParaInst(el,"GEO_Raggio") == None:
			param_inst_999.append(Para(el,"GEO_Raggio"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Raggio") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Raggio").Definition.Name))
		elif ParaInst(el,"GEO_Raggio") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Raggio"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Raggio"))
		except:
			pass

#-----GEO_Spessore (len__inst)
	if opera_el in ["CV","IM","MA","PV","SI"] and parteopera_el in ["AC","IC","IA","SP","IE","BA","BS"] and elemento_el in ["CAP","OPO","TIM","PPZ","SOL","SSB","MAN","MDA","MFR","PAR","CEE","QEB","QEM","BAN","PAN"]:
		if Para(el,"GEO_Spessore").HasValue == False or ParaInst(el,"GEO_Spessore") == None:
			param_inst_999.append(Para(el,"GEO_Spessore"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Spessore") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Spessore").Definition.Name))
		elif ParaInst(el,"GEO_Spessore") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Spessore"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Spessore"))
		except:
			pass

#-----GEO_Spessore parete (len__inst)
	if opera_el in ["CV","IM","PV"] and parteopera_el in ["FP","FO","IE"] and elemento_el in ["DIA","MPL","PAL","POZ","PZE"]:
		if Para(el,"GEO_Spessore parete").HasValue == False or ParaInst(el,"GEO_Spessore parete") == None:
			param_inst_999.append(Para(el,"GEO_Spessore parete"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Spessore parete") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Spessore parete").Definition.Name))
		elif ParaInst(el,"GEO_Spessore parete") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Spessore parete"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Spessore parete"))
		except:
			pass

#-----GEO_Spessore strati bituminosi (len__inst)
	if opera_el in ["CV","PV"] and parteopera_el in ["IA","IC"] and elemento_el in ["PPZ"]:
		if Para(el,"GEO_Spessore strati bituminosi").HasValue == False or ParaInst(el,"GEO_Spessore strati bituminosi") == None:
			param_inst_999.append(Para(el,"GEO_Spessore strati bituminosi"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Spessore strati bituminosi") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Spessore strati bituminosi").Definition.Name))
		elif ParaInst(el,"GEO_Spessore strati bituminosi") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Spessore strati bituminosi"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Spessore strati bituminosi"))
		except:
			pass


#-----GEO_Spessore strato pavimentazione (len__inst)
	if opera_el in ["CV","PV"] and parteopera_el in ["IA","IC"] and elemento_el in ["BIN","MUS"]:
		if Para(el,"GEO_Spessore strato pavimentazione").HasValue == False or ParaInst(el,"GEO_Spessore strato pavimentazione") == None:
			param_inst_999.append(Para(el,"GEO_Spessore strato pavimentazione"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Spessore strato pavimentazione") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Spessore strato pavimentazione").Definition.Name))
		elif ParaInst(el,"GEO_Spessore strato pavimentazione") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"GEO_Spessore strato pavimentazione"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Spessore strato pavimentazione"))
		except:
			pass

#-----GEO_Volume (Volume__inst)
	if opera_el in ["CV","IM","MA","PV"] and parteopera_el in ["AN","AC","FP","FO","FS","IC","IA","PD","PI","SP","IE","BA"] and elemento_el in ["SEL","TRV","SAR","DIA","MPL","PAL","PZF","POZ","CAS","COR","PPZ","SGE","SOL","SSB","TRA","BAG","PUL","MAN","MDA","MFR","PAR","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","DIE","IMT","INS","INM","PLI","REL","REP","RIF","SCS","SDE","TRS"]:
		if Para(el,"GEO_Volume").HasValue == False or ParaInst(el,"GEO_Volume") == None:
			param_inst_999.append(Para(el,"GEO_Volume"))
			inst_999.append(el)
		elif ParaInst(el,"GEO_Volume") == 35279.3520547671:
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"GEO_Volume").Definition.Name))
		elif Para(el,"GEO_Volume").AsValueString() == "111 m³":
			para_111_trasforma.append(Para(el,"GEO_Volume"))

	else:
		try:
			para_num_eccesso.append(Para(el,"GEO_Volume"))
		except:
			pass

#-------------------------------------IDENTIFICATIVO OGGETTO

#-----IDE_Codice opera (txt_inst)

	if opera_el in ["CV","IM","MA","MO","PV","SI"] and parteopera_el in ["AN","AC","FP","FO","FS","GI","IC","IA","PD","PI","SP","IE","BA","BE","CP","EA","ES","MN","TR","BS","BC"] and elemento_el in ["STL","SEL","TRV","CAP","OPO","PEN","SAR","TIM","DIA","MPL","PAL","PZF","CEN","POZ","GIU","AAP","BIN","CAS","COR","CUN","MUS","PPZ","SGE","SOL","SSB","TRA","VEL","CNT","LOR","PUN","SBL","TAN","AMS","BAG","ISA","PUL","RIS","MAN","MDA","MFR","PAR","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","CEE","DIE","IMT","INS","INM","MUL","PLI","PZE","QEB","QEM","REL","REP","RIF","SCS","SDE","TRS","BAN","BPT","MON","BRE","CLI","TSS","IFS","TEC","TCM","ACC","SCA","VCM","IDO","SEM","SUM","TIG","LMC","PAN","RAN","TTA","UNI","NJE"]:
		if Para(el,"IDE_Codice opera").HasValue == False or ParaInst(el,"IDE_Codice opera") == "" or ParaInst(el,"IDE_Codice opera") == None:
			param_inst_compilare.append(Para(el,"IDE_Codice opera"))
			inst_compilare.append(el)
		elif ParaInst(el,"IDE_Codice opera") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"IDE_Codice opera").Definition.Name))
		elif ParaInst(el,"IDE_Codice opera") == "ND":
			para_ND_trasforma.append(Para(el,"IDE_Codice opera"))

	else:
		try:
			para_eccesso.append(Para(el,"IDE_Codice opera"))
		except:
			pass

#-----IDE_Codice sensore (txt_inst)

	if opera_el in ["MO"] and parteopera_el in ["BE","CP","EA","ES","MN","TR"] and elemento_el in ["BRE","CLI","TSS","IFS","TEC","TCM","ACC","SCA","VCM","IDO","SEM","SUM","TIG"]:
		if Para(el,"IDE_Codice sensore").HasValue == False or ParaInst(el,"IDE_Codice sensore") == "" or ParaInst(el,"IDE_Codice sensore") == None:
			param_inst_compilare.append(Para(el,"IDE_Codice sensore"))
			inst_compilare.append(el)
		elif ParaInst(el,"IDE_Codice sensore") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"IDE_Codice sensore").Definition.Name))
		elif ParaInst(el,"IDE_Codice sensore") == "ND":
			para_ND_trasforma.append(Para(el,"IDE_Codice sensore"))

	else:
		try:
			para_eccesso.append(Para(el,"IDE_Codice sensore"))
		except:
			pass

#-----IDE_Codice WBS (txt_inst)

	if opera_el in ["CV","IM","MA","MO","PV","SI"] and parteopera_el in ["AN","AC","FP","FO","FS","GI","IC","IA","PD","PI","SP","IE","BA","BE","CP","EA","ES","MN","TR","BS","BC"] and elemento_el in ["STL","SEL","TRV","CAP","OPO","PEN","SAR","TIM","DIA","MPL","PAL","PZF","CEN","POZ","GIU","AAP","BIN","CAS","COR","CUN","MUS","PPZ","SGE","SOL","SSB","TRA","VEL","CNT","LOR","PUN","SBL","TAN","AMS","BAG","ISA","PUL","RIS","MAN","MDA","MFR","PAR","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","CEE","DIE","IMT","INS","INM","MUL","PLI","PZE","QEB","QEM","REL","REP","RIF","SCS","SDE","TRS","BAN","BPT","MON","BRE","CLI","TSS","IFS","TEC","TCM","ACC","SCA","VCM","IDO","SEM","SUM","TIG","LMC","PAN","RAN","TTA","UNI","NJE"]:
		if Para(el,"IDE_Codice WBS").HasValue == False or ParaInst(el,"IDE_Codice WBS") == "" or ParaInst(el,"IDE_Codice WBS") == None:
			param_inst_compilare.append(Para(el,"IDE_Codice WBS"))
			inst_compilare.append(el)
		elif ParaInst(el,"IDE_Codice WBS") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"IDE_Codice WBS").Definition.Name))
		elif ParaInst(el,"IDE_Codice WBS") == "ND":
			para_ND_trasforma.append(Para(el,"IDE_Codice WBS"))

	else:
		try:
			para_eccesso.append(Para(el,"IDE_Codice WBS"))
		except:
			pass

#-----IDE_Elemento di appartenenza (txt_inst)

	if opera_el in ["CV","MO","PV"] and parteopera_el in ["IC","IA","BE","CP","EA","ES","MN","TR"] and elemento_el in ["SGE","BRE","CLI","TSS","IFS","TEC","TCM","ACC","SCA","VCM","IDO","SEM","SUM","TIG"]:
		if Para(el,"IDE_Elemento di appartenenza").HasValue == False or ParaInst(el,"IDE_Elemento di appartenenza") == "" or ParaInst(el,"IDE_Elemento di appartenenza") == None:
			param_inst_compilare.append(Para(el,"IDE_Elemento di appartenenza"))
			inst_compilare.append(el)
		elif ParaInst(el,"IDE_Elemento di appartenenza") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"IDE_Elemento di appartenenza").Definition.Name))
		elif ParaInst(el,"IDE_Elemento di appartenenza") == "ND":
			para_ND_trasforma.append(Para(el,"IDE_Elemento di appartenenza"))

	else:
		try:
			para_eccesso.append(Para(el,"IDE_Elemento di appartenenza"))
		except:
			pass

#-----IDE_Gruppo anagrafica (txt_inst)

	if opera_el in ["CV","IM","MA","MO","PV","SI"] and parteopera_el in ["AN","AC","FP","FO","FS","GI","IC","IA","PD","PI","SP","IE","BA","BE","CP","EA","ES","MN","TR","BS","BC"] and elemento_el in ["STL","SEL","TRV","CAP","OPO","PEN","SAR","TIM","DIA","MPL","PAL","PZF","CEN","POZ","GIU","AAP","BIN","CAS","COR","CUN","MUS","PPZ","SGE","SOL","SSB","TRA","VEL","CNT","LOR","PUN","SBL","TAN","AMS","BAG","ISA","PUL","RIS","MAN","MDA","MFR","PAR","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","CEE","DIE","IMT","INS","INM","MUL","PLI","PZE","QEB","QEM","REL","REP","RIF","SCS","SDE","TRS","BAN","BPT","MON","BRE","CLI","TSS","IFS","TEC","TCM","ACC","SCA","VCM","IDO","SEM","SUM","TIG","LMC","PAN","RAN","TTA","UNI","NJE"]:
		if Para(el,"IDE_Gruppo anagrafica").HasValue == False or ParaInst(el,"IDE_Gruppo anagrafica") == "" or ParaInst(el,"IDE_Gruppo anagrafica") == None:
			param_inst_compilare.append(Para(el,"IDE_Gruppo anagrafica"))
			inst_compilare.append(el)
		elif ParaInst(el,"IDE_Gruppo anagrafica") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"IDE_Gruppo anagrafica").Definition.Name))
		elif ParaInst(el,"IDE_Gruppo anagrafica") == "ND":
			para_ND_trasforma.append(Para(el,"IDE_Gruppo anagrafica"))

	else:
		try:
			para_eccesso.append(Para(el,"IDE_Gruppo anagrafica"))
		except:
			pass

#-----IDE_LOR (txt_inst)

	if opera_el in ["CV","IM","MA","MO","PV","SI"] and parteopera_el in ["AN","AC","FP","FO","FS","GI","IC","IA","PD","PI","SP","IE","BA","BE","CP","EA","ES","MN","TR","BS","BC"] and elemento_el in ["STL","SEL","TRV","CAP","OPO","PEN","SAR","TIM","DIA","MPL","PAL","PZF","CEN","POZ","GIU","AAP","BIN","CAS","COR","CUN","MUS","PPZ","SGE","SOL","SSB","TRA","VEL","CNT","LOR","PUN","SBL","TAN","AMS","BAG","ISA","PUL","RIS","MAN","MDA","MFR","PAR","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","CEE","DIE","IMT","INS","INM","MUL","PLI","PZE","QEB","QEM","REL","REP","RIF","SCS","SDE","TRS","BAN","BPT","MON","BRE","CLI","TSS","IFS","TEC","TCM","ACC","SCA","VCM","IDO","SEM","SUM","TIG","LMC","PAN","RAN","TTA","UNI","NJE"]:
		if Para(el,"IDE_LOR").HasValue == False or ParaInst(el,"IDE_LOR") == "" or ParaInst(el,"IDE_LOR") == None:
			param_inst_compilare.append(Para(el,"IDE_LOR"))
			inst_compilare.append(el)
		elif ParaInst(el,"IDE_LOR") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"IDE_LOR").Definition.Name))
		elif ParaInst(el,"IDE_LOR") == "ND":
			para_ND_trasforma.append(Para(el,"IDE_LOR"))

	else:
		try:
			para_eccesso.append(Para(el,"IDE_LOR"))
		except:
			pass

#-----IDE_Prova eseguita su: (txt_inst)

	if opera_el in ["MO"] and parteopera_el in ["ES"] and elemento_el in ["TCM"]:
		if Para(el,"IDE_Prova eseguita su:").HasValue == False or ParaInst(el,"IDE_Prova eseguita su:") == "" or ParaInst(el,"IDE_Prova eseguita su:") == None:
			param_inst_compilare.append(Para(el,"IDE_Prova eseguita su:"))
			inst_compilare.append(el)
		elif ParaInst(el,"IDE_Prova eseguita su:") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"IDE_Prova eseguita su:").Definition.Name))
		elif ParaInst(el,"IDE_Prova eseguita su:") == "ND":
			para_ND_trasforma.append(Para(el,"IDE_Prova eseguita su:"))

	else:
		try:
			para_eccesso.append(Para(el,"IDE_Prova eseguita su:"))
		except:
			pass

#-----IDE_Uniclass (txt_inst)

	if opera_el in ["CV","IM","MA","MO","PV","SI"] and parteopera_el in ["AN","AC","FP","FO","FS","GI","IC","IA","PD","PI","SP","IE","BA","BE","CP","EA","ES","MN","TR","BS","BC"] and elemento_el in ["STL","SEL","TRV","CAP","OPO","PEN","SAR","TIM","DIA","MPL","PAL","PZF","CEN","POZ","GIU","AAP","BIN","CAS","COR","CUN","MUS","PPZ","SGE","SOL","SSB","TRA","VEL","CNT","LOR","PUN","SBL","TAN","AMS","BAG","ISA","PUL","RIS","MAN","MDA","MFR","PAR","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","CEE","DIE","IMT","INS","INM","MUL","PLI","PZE","QEB","QEM","REL","REP","RIF","SCS","SDE","TRS","BAN","BPT","MON","BRE","CLI","TSS","IFS","TEC","TCM","ACC","SCA","VCM","IDO","SEM","SUM","TIG","LMC","PAN","RAN","TTA","UNI","NJE"]:
		if Para(el,"IDE_Uniclass").HasValue == False or ParaInst(el,"IDE_Uniclass") == "" or ParaInst(el,"IDE_Uniclass") == None:
			param_inst_compilare.append(Para(el,"IDE_Uniclass"))
			inst_compilare.append(el)
		elif ParaInst(el,"IDE_Uniclass") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"IDE_Uniclass").Definition.Name))
		elif ParaInst(el,"IDE_Uniclass") == "ND":
			para_ND_trasforma.append(Para(el,"IDE_Uniclass"))

	else:
		try:
			para_eccesso.append(Para(el,"IDE_Uniclass"))
		except:
			pass

#-------------------------------------INFORMAZIONI 6D

#-----INF_Campata di appartenenza (txt_inst)
#-----INF_Codice BMS (txt_inst)
#-----INF_Impalcato di appartenenza (txt_inst)
#-----INF_Numerazione struttura campata (txt_inst)

	if opera_el in ["CV","PV"] and parteopera_el in ["AC","IC","IA","PD"] and elemento_el in ["PEN","SAR","TRV","AAP","BIN","CAS","COR","CUN","MUS","PPZ","SGE","SOL","SSB","TRA","VEL","CNT","LOR","PUN","SBL","TAN","AMS","BAG","ISA","PUL","RIS","SEL"]:
		if Para(el,"INF_Campata di appartenenza").HasValue == False or ParaInst(el,"INF_Campata di appartenenza") == "" or ParaInst(el,"INF_Campata di appartenenza") == None:
			param_inst_compilare.append(Para(el,"INF_Campata di appartenenza"))
			inst_compilare.append(el)
		elif ParaInst(el,"INF_Campata di appartenenza") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"INF_Campata di appartenenza").Definition.Name))
		elif ParaInst(el,"INF_Campata di appartenenza") == "ND":
			para_ND_trasforma.append(Para(el,"INF_Campata di appartenenza"))


		if Para(el,"INF_Codice BMS").HasValue == False or ParaInst(el,"INF_Codice BMS") == "" or ParaInst(el,"INF_Codice BMS") == None:
			param_inst_compilare.append(Para(el,"INF_Codice BMS"))
			inst_compilare.append(el)
		elif ParaInst(el,"INF_Codice BMS") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"INF_Codice BMS").Definition.Name))
		elif ParaInst(el,"INF_Codice BMS") == "ND":
			para_ND_trasforma.append(Para(el,"INF_Codice BMS"))

			
		if Para(el,"INF_Impalcato di appartenenza").HasValue == False or ParaInst(el,"INF_Impalcato di appartenenza") == "" or ParaInst(el,"INF_Impalcato di appartenenza") == None:
			param_inst_compilare.append(Para(el,"INF_Impalcato di appartenenza"))
			inst_compilare.append(el)
		elif ParaInst(el,"INF_Impalcato di appartenenza") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"INF_Impalcato di appartenenza").Definition.Name))
		elif ParaInst(el,"INF_Impalcato di appartenenza") == "ND":
			para_ND_trasforma.append(Para(el,"INF_Impalcato di appartenenza"))


		if Para(el,"INF_Numerazione struttura campata").HasValue == False or ParaInst(el,"INF_Numerazione struttura campata") == "" or ParaInst(el,"INF_Numerazione struttura campata") == None:
			param_inst_compilare.append(Para(el,"INF_Numerazione struttura campata"))
			inst_compilare.append(el)
		elif ParaInst(el,"INF_Numerazione struttura campata") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"INF_Numerazione struttura campata").Definition.Name))
		elif ParaInst(el,"INF_Numerazione struttura campata") == "ND":
			para_ND_trasforma.append(Para(el,"INF_Numerazione struttura campata"))

	else:
		try:
			para_eccesso.append(Para(el,"INF_Campata di appartenenza"))
			para_eccesso.append(Para(el,"INF_Codice BMS"))
			para_eccesso.append(Para(el,"INF_Impalcato di appartenenza"))
			para_eccesso.append(Para(el,"INF_Numerazione struttura campata"))
		except:
			pass

#-------------------------------------LOCALIZZAZIONE

#-----LOC_Carreggiata (txt_inst)
#-----LOC_Direzione (txt_inst)

	if opera_el in ["CV","IM","MA","PV","SI"] and parteopera_el in ["AN","AC","FP","FO","FS","GI","IC","IA","PD","PI","SP","IE","BA","BS","BC"] and elemento_el in ["STL","SEL","TRV","CAP","OPO","PEN","SAR","TIM","DIA","MPL","PAL","PZF","CEN","POZ","GIU","AAP","BIN","CAS","COR","CUN","MUS","PPZ","SGE","SOL","SSB","TRA","VEL","CNT","LOR","PUN","SBL","TAN","AMS","BAG","ISA","PUL","RIS","MAN","MDA","MFR","PAR","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","CEE","DIE","IMT","INS","INM","MUL","PLI","PZE","QEB","QEM","REL","REP","RIF","SCS","SDE","TRS","BAN","BPT","MON","LMC","PAN","RAN","TTA","UNI","NJE"]:

		if Para(el,"LOC_Carreggiata").HasValue == False or ParaInst(el,"LOC_Carreggiata") == "" or ParaInst(el,"LOC_Carreggiata") == None:
			param_inst_compilare.append(Para(el,"LOC_Carreggiata"))
			inst_compilare.append(el)
		elif ParaInst(el,"LOC_Carreggiata") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"LOC_Carreggiata").Definition.Name))
		elif ParaInst(el,"LOC_Carreggiata") == "ND":
			para_ND_trasforma.append(Para(el,"LOC_Carreggiata"))


		if Para(el,"LOC_Direzione").HasValue == False or ParaInst(el,"LOC_Direzione") == "" or ParaInst(el,"LOC_Direzione") == None:
			param_inst_compilare.append(Para(el,"LOC_Direzione"))
			inst_compilare.append(el)
		elif ParaInst(el,"LOC_Direzione") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"LOC_Direzione").Definition.Name))
		elif ParaInst(el,"LOC_Direzione") == "ND":
			para_ND_trasforma.append(Para(el,"LOC_Direzione"))

			
	else:
		try:
			para_eccesso.append(Para(el,"LOC_Carreggiata"))
			para_eccesso.append(Para(el,"LOC_Direzione"))
		except:
			pass

#-----LOC_Progressiva finale (txt_inst)
#-----LOC_Progressiva iniziale (txt_inst)

	if opera_el in ["CV","IM","MA","PV","SI"] and parteopera_el in ["AN","AC","IC","IA","PD","PI","SP","IE","BA","BS","BC"] and elemento_el in ["SEL","CAP","OPO","TIM","AAP","MUS","AMS","ISA","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","CEE","DIE","IMT","INS","INM","MUL","PZE","QEB","QEM","REL","REP","RIF","SCS","SDE","TRS","BAN","LMC","PAN","RAN","TTA","NJE"]:

		if Para(el,"LOC_Progressiva finale").HasValue == False or ParaInst(el,"LOC_Progressiva finale") == "" or ParaInst(el,"LOC_Progressiva finale") == None:
			param_inst_compilare.append(Para(el,"LOC_Progressiva finale"))
			inst_compilare.append(el)
		elif ParaInst(el,"LOC_Progressiva finale") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"LOC_Progressiva finale").Definition.Name))
		elif ParaInst(el,"LOC_Progressiva finale") == "ND":
			para_ND_trasforma.append(Para(el,"LOC_Progressiva finale"))

		
		if Para(el,"LOC_Progressiva iniziale").HasValue == False or ParaInst(el,"LOC_Progressiva iniziale") == "" or ParaInst(el,"LOC_Progressiva iniziale") == None:
			param_inst_compilare.append(Para(el,"LOC_Progressiva iniziale"))
			inst_compilare.append(el)
		elif ParaInst(el,"LOC_Progressiva iniziale") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"LOC_Progressiva iniziale").Definition.Name))
		elif ParaInst(el,"LOC_Progressiva iniziale") == "ND":
			para_ND_trasforma.append(Para(el,"LOC_Progressiva iniziale"))

			
	else:
		try:
			para_eccesso.append(Para(el,"LOC_Progressiva finale"))
			para_eccesso.append(Para(el,"LOC_Progressiva iniziale"))
		except:
			pass

#-------------------------------------TECNICO

#-----TEC_Chiave di taglio (txt_inst)

	if opera_el in ["CV","PV"] and parteopera_el in ["IC","IA"] and elemento_el in ["CAS","SGE","TRA"]:

		if Para(el,"TEC_Chiave di taglio").HasValue == False or ParaInst(el,"TEC_Chiave di taglio") == "" or ParaInst(el,"TEC_Chiave di taglio") == None:
			param_inst_compilare.append(Para(el,"TEC_Chiave di taglio"))
			inst_compilare.append(el)
		elif ParaInst(el,"TEC_Chiave di taglio") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"TEC_Chiave di taglio").Definition.Name))
		elif ParaInst(el,"TEC_Chiave di taglio") == "ND":
			para_ND_trasforma.append(Para(el,"TEC_Chiave di taglio"))

	else:
		try:
			para_eccesso.append(Para(el,"TEC_Chiave di taglio"))
		except:
			pass

#-----TEC_Classe di prestazione (txt_inst)

	if opera_el in ["SI"] and parteopera_el in ["BC","BS"] and elemento_el in ["LMC","TTA","NJE"]:

		if Para(el,"TEC_Classe di prestazione").HasValue == False or ParaInst(el,"TEC_Classe di prestazione") == "" or ParaInst(el,"TEC_Classe di prestazione") == None:
			param_inst_compilare.append(Para(el,"TEC_Classe di prestazione"))
			inst_compilare.append(el)
		elif ParaInst(el,"TEC_Classe di prestazione") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"TEC_Classe di prestazione").Definition.Name))
		elif ParaInst(el,"TEC_Classe di prestazione") == "ND":
			para_ND_trasforma.append(Para(el,"TEC_Classe di prestazione"))

	else:
		try:
			para_eccesso.append(Para(el,"TEC_Classe di prestazione"))
		except:
			pass

#-----TEC_Dimensione maglia (area__inst)
	if opera_el in ["SI"] and parteopera_el in ["BS"] and elemento_el in ["RAN"]:
		if Para(el,"TEC_Dimensione maglia").HasValue == False or ParaInst(el,"TEC_Dimensione maglia") == None:
			param_inst_999.append(Para(el,"TEC_Dimensione maglia"))
			inst_999.append(el)
		elif ParaInst(el,"TEC_Dimensione maglia") == ConvUnitsFqMq(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"TEC_Dimensione maglia").Definition.Name))
		elif Para(el,"TEC_Dimensione maglia").AsValueString() == "111 m²":
			para_111_trasforma.append(Para(el,"TEC_Dimensione maglia"))
	else:
		try:
			para_num_eccesso.append(Para(el,"TEC_Dimensione maglia"))
		except:
			pass

#-----TEC_Incidenza armatura (mass__inst)
	if opera_el in ["CV","IM","MA","PV"] and parteopera_el in ["AN","AC","FP","FO","FS","IC","IA","PD","PI","SP","IE","BA"] and elemento_el in ["SEL","TRV","SAR","DIA","MPL","PAL","PZF","CAS","COR","SGE","SOL","SSB","TRA","BAG","PUL","MAN","MDA","MFR","PAR","PLI"]:
		if Para(el,"TEC_Incidenza armatura").HasValue == False or ParaInst(el,"TEC_Incidenza armatura") == None:
			param_inst_999.append(Para(el,"TEC_Incidenza armatura"))
			inst_999.append(el)
		elif ParaInst(el,"TEC_Incidenza armatura") == 28.288529745408:
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"TEC_Incidenza armatura").Definition.Name))
		elif Para(el,"TEC_Incidenza armatura").AsValueString() == "111.00 kg/m³":
			para_111_trasforma.append(Para(el,"TEC_Incidenza armatura"))
	else:
		try:
			para_num_eccesso.append(Para(el,"TEC_Incidenza armatura"))
		except:
			pass
			

#-----TEC_Livello idrico (len__inst)
	if opera_el in ["CV","PV"] and parteopera_el in ["SP"] and elemento_el in ["MFR","PAR"]:
		if Para(el,"TEC_Livello idrico").HasValue == False or ParaInst(el,"TEC_Livello idrico") == None:
			param_inst_999.append(Para(el,"TEC_Livello idrico"))
			inst_999.append(el)
		elif ParaInst(el,"TEC_Livello idrico") == ConvUnitsFM(999):
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"TEC_Livello idrico").Definition.Name))
		elif ParaInst(el,"TEC_Livello idrico") == ConvUnitsFM(111):
			para_111_trasforma.append(Para(el,"TEC_Livello idrico"))

	else:
		try:
			para_num_eccesso.append(Para(el,"TEC_Livello idrico"))
		except:
			pass

#-----TEC_Documentazione tecnica (txt_inst)

	if opera_el in ["CV","IM","MA","MO","PV","SI"] and parteopera_el in ["AN","AC","FP","FO","FS","GI","IC","IA","PD","PI","SP","IE","BA","BE","CP","EA","ES","MN","TR","BS","BC"] and elemento_el in ["STL","SEL","TRV","CAP","OPO","PEN","SAR","TIM","DIA","MPL","PAL","PZF","CEN","POZ","GIU","AAP","BIN","CAS","COR","CUN","MUS","PPZ","SGE","SOL","SSB","TRA","VEL","CNT","LOR","PUN","SBL","TAN","AMS","BAG","ISA","PUL","RIS","MAN","MDA","MFR","PAR","ALI","BLI","CAB","CNP","CAE","CAV","COL","COM","CEE","DIE","IMT","INS","INM","MUL","PLI","PZE","QEB","QEM","REL","REP","RIF","SCS","SDE","TRS","BAN","BPT","MON","BRE","CLI","TSS","IFS","TEC","TCM","ACC","SCA","VCM","IDO","SEM","SUM","TIG","LMC","PAN","RAN","TTA","UNI","NJE"]:

		if Para(el,"TEC_Documentazione tecnica").HasValue == False or ParaInst(el,"TEC_Documentazione tecnica") == "" or ParaInst(el,"TEC_Documentazione tecnica") == None:
			param_inst_compilare.append(Para(el,"TEC_Documentazione tecnica"))
			inst_compilare.append(el)
		elif ParaInst(el,"TEC_Documentazione tecnica") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"TEC_Documentazione tecnica").Definition.Name))
		elif ParaInst(el,"TEC_Documentazione tecnica") == "ND":
			para_ND_trasforma.append(Para(el,"TEC_Documentazione tecnica"))

	else:
		try:
			para_eccesso.append(Para(el,"TEC_Documentazione tecnica"))
		except:
			pass

#-----TEC_Numero seriale (txt_inst)

	if opera_el in ["MO"] and parteopera_el in ["BE","CP","EA","ES","MN","TR"] and elemento_el in ["BRE","CLI","TSS","IFS","TEC","TCM","ACC","SCA","VCM","IDO","SEM","SUM","TIG"]:
		if Para(el,"TEC_Numero seriale").HasValue == False or ParaInst(el,"TEC_Numero seriale") == "" or ParaInst(el,"TEC_Numero seriale") == None:
			param_inst_compilare.append(Para(el,"TEC_Numero seriale"))
			inst_compilare.append(el)
		elif ParaInst(el,"TEC_Numero seriale") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"TEC_Numero seriale").Definition.Name))
		elif ParaInst(el,"TEC_Numero seriale") == "ND":
			para_ND_trasforma.append(Para(el,"TEC_Numero seriale"))

	else:
		try:
			para_eccesso.append(Para(el,"TEC_Numero seriale"))
		except:
			pass

#-----TEC_Posizione (txt_inst)

	if opera_el in ["MO"] and parteopera_el in ["ES","MN","TR"] and elemento_el in ["IFS","TEC","TCM","ACC","SCA","VCM","SEM","SUM","TIG"]:
		if Para(el,"TEC_Posizione").HasValue == False or ParaInst(el,"TEC_Posizione") == "" or ParaInst(el,"TEC_Posizione") == None:
			param_inst_compilare.append(Para(el,"TEC_Posizione"))
			inst_compilare.append(el)
		elif ParaInst(el,"TEC_Posizione") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"TEC_Posizione").Definition.Name))
		elif ParaInst(el,"TEC_Posizione") == "ND":
			para_ND_trasforma.append(Para(el,"TEC_Posizione"))
	else:
		try:
			para_eccesso.append(Para(el,"TEC_Posizione"))
		except:
			pass

#-----TEC_Sella Gerber (txt_inst)

	if opera_el in ["CV","PV"] and parteopera_el in ["IC","IA","PI"] and elemento_el in ["CAS","PUL","TRA"]:

		if Para(el,"TEC_Sella Gerber").HasValue == False or ParaInst(el,"TEC_Sella Gerber") == "" or ParaInst(el,"TEC_Sella Gerber") == None:
			param_inst_compilare.append(Para(el,"TEC_Sella Gerber"))
			inst_compilare.append(el)
		elif ParaInst(el,"TEC_Sella Gerber") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"TEC_Sella Gerber").Definition.Name))
		elif ParaInst(el,"TEC_Sella Gerber") == "ND":
			para_ND_trasforma.append(Para(el,"TEC_Sella Gerber"))
	else:
		try:
			para_eccesso.append(Para(type_el,"TEC_Sella Gerber"))
		except:
			pass

#-----TEC_Tipologia (txt_type)

	if opera_el in ["CV","MA","PV"] and parteopera_el in ["AN","GI","IC","IA","SP","BA"] and elemento_el in ["STL","GIU","SOL","BAN"]:

		if Para(type_el,"TEC_Tipologia").HasValue == False or ParaType(el,"TEC_Tipologia",doc) == "" or ParaType(el,"TEC_Tipologia",doc) == None:
			param_inst_compilare.append(Para(type_el,"TEC_Tipologia"))
			inst_compilare.append(type_el)
		elif ParaType(el,"TEC_Tipologia",doc) == "•••COMPILARE•••":
			para_typ_check.add("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(type_el.Id),Para(type_el,"TEC_Tipologia").Definition.Name))
		elif ParaType(el,"TEC_Tipologia",doc) == "ND":
			para_ND_trasforma.append(Para(type_el,"TEC_Tipologia"))

	else:
		try:
			para_eccesso.append(Para(type_el,"TEC_Tipologia"))
		except:
			pass

#-----TEC_Tipologia acciaio (txt_type)

	if opera_el in ["CV","IM","MA","PV","SI"] and parteopera_el in ["AN","AC","FP","FO","FS","IC","IA","PD","PI","SP","IE","BA","BS","BC"] and elemento_el in ["SEL","TRV","SAR","DIA","MPL","PAL","PZF","POZ","CAS","COR","SGE","SOL","SSB","TRA","CNT","LOR","PUN","SBL","TAN","BAG","PUL","RIS","MAN","MDA","MFR","PAR","PLI","BAN","MON","PAN","RAN","UNI"]:

		if Para(type_el,"TEC_Tipologia acciaio").HasValue == False or ParaType(el,"TEC_Tipologia acciaio",doc) == "" or ParaType(el,"TEC_Tipologia acciaio",doc) == None:
			param_inst_compilare.append(Para(type_el,"TEC_Tipologia acciaio"))
			inst_compilare.append(type_el)
		elif ParaType(el,"TEC_Tipologia acciaio",doc) == "•••COMPILARE•••":
			para_typ_check.add("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(type_el.Id),Para(type_el,"TEC_Tipologia acciaio").Definition.Name))
		elif ParaType(el,"TEC_Tipologia acciaio",doc) == "ND":
			para_ND_trasforma.append(Para(type_el,"TEC_Tipologia acciaio"))
	else:
		try:
			para_eccesso.append(Para(type_el,"TEC_Tipologia acciaio"))
		except:
			pass

#-----TEC_Tipologia apparecchi di appoggio (txt_type)

#-----TEC_Tipologia dispositivo antisismico (txt_type)

	if opera_el in ["CV","PV"] and parteopera_el in ["PD","PI","SP"] and elemento_el in ["ISA"]:

		if Para(type_el,"TEC_Tipologia apparecchi di appoggio").HasValue == False or ParaType(el,"TEC_Tipologia apparecchi di appoggio",doc) == "" or ParaType(el,"TEC_Tipologia apparecchi di appoggio",doc) == None:
			param_inst_compilare.append(Para(type_el,"TEC_Tipologia apparecchi di appoggio"))
			inst_compilare.append(type_el)
		elif ParaType(el,"TEC_Tipologia apparecchi di appoggio",doc) == "•••COMPILARE•••":
			para_typ_check.add("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(type_el.Id),Para(type_el,"TEC_Tipologia apparecchi di appoggio").Definition.Name))
		elif ParaType(el,"TEC_Tipologia apparecchi di appoggio",doc) == "ND":
			para_ND_trasforma.append(Para(type_el,"TEC_Tipologia apparecchi di appoggio"))

		if Para(type_el,"TEC_Tipologia dispositivo antisismico").HasValue == False or ParaType(el,"TEC_Tipologia dispositivo antisismico",doc) == "" or ParaType(el,"TEC_Tipologia dispositivo antisismico",doc) == None:
			param_inst_compilare.append(Para(type_el,"TEC_Tipologia dispositivo antisismico"))
			inst_compilare.append(type_el)
		elif ParaType(el,"TEC_Tipologia dispositivo antisismico",doc) == "•••COMPILARE•••":
			para_typ_check.add("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(type_el.Id),Para(type_el,"TEC_Tipologia dispositivo antisismico").Definition.Name))
		elif ParaType(el,"TEC_Tipologia dispositivo antisismico",doc) == "ND":
			para_ND_trasforma.append(Para(type_el,"TEC_Tipologia dispositivo antisismico"))
	else:
		try:
			para_eccesso.append(Para(type_el,"TEC_Tipologia apparecchi di appoggio"))
			para_eccesso.append(Para(type_el,"TEC_Tipologia dispositivo antisismico"))
		except:
			pass

#-----TEC_Tipologia barriera (txt_type)

	if opera_el in ["SI"] and parteopera_el in ["BS"] and elemento_el in ["LMC"]:

		if Para(type_el,"TEC_Tipologia barriera").HasValue == False or ParaType(el,"TEC_Tipologia barriera",doc) == "" or ParaType(el,"TEC_Tipologia barriera",doc) == None:
			param_inst_compilare.append(Para(type_el,"TEC_Tipologia barriera"))
			inst_compilare.append(type_el)
		elif ParaType(el,"TEC_Tipologia barriera",doc) == "•••COMPILARE•••":
			para_typ_check.add("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(type_el.Id),Para(type_el,"TEC_Tipologia barriera").Definition.Name))
		elif ParaType(el,"TEC_Tipologia barriera",doc) == "ND":
			para_ND_trasforma.append(Para(type_el,"TEC_Tipologia barriera"))
	else:
		try:
			para_eccesso.append(Para(type_el,"TEC_Tipologia barriera"))
		except:
			pass

#-----TEC_Tipologia C.A./C.A.P. (txt_type)

	if opera_el in ["CV","PV"] and parteopera_el in ["AN","AC","IC","IA","PD","PI"] and elemento_el in ["TRV","CAS","SGE","TRA"]:
		if Para(type_el,"TEC_Tipologia C.A./C.A.P.").HasValue == False or ParaType(el,"TEC_Tipologia C.A./C.A.P.",doc) == "" or ParaType(el,"TEC_Tipologia C.A./C.A.P.",doc) == None:
			param_inst_compilare.append(Para(type_el,"TEC_Tipologia C.A./C.A.P."))
			inst_compilare.append(type_el)
		elif ParaType(el,"TEC_Tipologia C.A./C.A.P.",doc) == "•••COMPILARE•••":
			para_typ_check.add("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(type_el.Id),Para(type_el,"TEC_Tipologia C.A./C.A.P.").Definition.Name))
		elif ParaType(el,"TEC_Tipologia C.A./C.A.P.",doc) == "ND":
			para_ND_trasforma.append(Para(type_el,"TEC_Tipologia C.A./C.A.P."))
	else:
		try:
			para_eccesso.append(Para(type_el,"TEC_Tipologia C.A./C.A.P."))
		except:
			pass

#-----TEC_Tipologia finitura (txt_type)

	if opera_el in ["MA"] and parteopera_el in ["BA"] and elemento_el in ["MON"]:

		if Para(type_el,"TEC_Tipologia finitura").HasValue == False or ParaType(el,"TEC_Tipologia finitura",doc) == "" or ParaType(el,"TEC_Tipologia finitura",doc) == None:
			param_inst_compilare.append(Para(type_el,"TEC_Tipologia finitura"))
			inst_compilare.append(type_el)
		elif ParaType(el,"TEC_Tipologia finitura",doc) == "•••COMPILARE•••":
			para_typ_check.add("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(type_el.Id),Para(type_el,"TEC_Tipologia finitura").Definition.Name))
		elif ParaType(el,"TEC_Tipologia finitura",doc) == "ND":
			para_ND_trasforma.append(Para(type_el,"TEC_Tipologia finitura"))
	else:
		try:
			para_eccesso.append(Para(type_el,"TEC_Tipologia finitura"))
		except:
			pass

#-----TEC_Tipologia giunti di superficie (txt_type)

	if opera_el in ["CV","PV"] and parteopera_el in ["GI"] and elemento_el in ["GIU"]:

		if Para(type_el,"TEC_Tipologia giunti di superficie").HasValue == False or ParaType(el,"TEC_Tipologia giunti di superficie",doc) == "" or ParaType(el,"TEC_Tipologia giunti di superficie",doc) == None:
			param_inst_compilare.append(Para(type_el,"TEC_Tipologia giunti di superficie"))
			inst_compilare.append(type_el)
		elif ParaType(el,"TEC_Tipologia giunti di superficie",doc) == "•••COMPILARE•••":
			para_typ_check.add("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(type_el.Id),Para(type_el,"TEC_Tipologia giunti di superficie").Definition.Name))
		elif ParaType(el,"TEC_Tipologia giunti di superficie",doc) == "ND":
			para_ND_trasforma.append(Para(type_el,"TEC_Tipologia giunti di superficie"))
	else:
		try:
			para_eccesso.append(Para(type_el,"TEC_Tipologia giunti di superficie"))
		except:
			pass

#-----TEC_Tipologia installazione (txt_type)

	if opera_el in ["CV","PV","SI"] and parteopera_el in ["AN","AC","IC","IA","BC","BS"] and elemento_el in ["STL","PEN","LMC","TTA","NJE"]:

		if Para(type_el,"TEC_Tipologia installazione").HasValue == False or ParaType(el,"TEC_Tipologia installazione",doc) == "" or ParaType(el,"TEC_Tipologia installazione",doc) == None:
			param_inst_compilare.append(Para(type_el,"TEC_Tipologia installazione"))
			inst_compilare.append(type_el)
		elif ParaType(el,"TEC_Tipologia installazione",doc) == "•••COMPILARE•••":
			para_typ_check.add("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(type_el.Id),Para(type_el,"TEC_Tipologia installazione").Definition.Name))
		elif ParaType(el,"TEC_Tipologia installazione",doc) == "ND":
			para_ND_trasforma.append(Para(type_el,"TEC_Tipologia installazione"))
	else:
		try:
			para_eccesso.append(Para(type_el,"TEC_Tipologia installazione"))
		except:
			pass

#-----TEC_Tipologia pendini (txt_type)

	if opera_el in ["CV","PV"] and parteopera_el in ["AC","IC","IA"] and elemento_el in ["PEN"]:

		if Para(type_el,"TEC_Tipologia pendini").HasValue == False or ParaType(el,"TEC_Tipologia pendini",doc) == "" or ParaType(el,"TEC_Tipologia pendini",doc) == None:
			param_inst_compilare.append(Para(type_el,"TEC_Tipologia pendini"))
			inst_compilare.append(type_el)
		elif ParaType(el,"TEC_Tipologia pendini",doc) == "•••COMPILARE•••":
			para_typ_check.add("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(type_el.Id),Para(type_el,"TEC_Tipologia pendini").Definition.Name))
		elif ParaType(el,"TEC_Tipologia pendini",doc) == "ND":
			para_ND_trasforma.append(Para(type_el,"TEC_Tipologia pendini"))
	else:
		try:
			para_eccesso.append(Para(type_el,"TEC_Tipologia pendini"))
		except:
			pass

#-----TEC_Tipologia profilo (txt_type)

	if opera_el in ["CV","PV","MA"] and parteopera_el in ["AN","AC","IC","IA","PD","PI","SP","BA"] and elemento_el in ["TRV","CNT","LOR","PUN","SBL","TAN","RIS","MON"]:

		if Para(type_el,"TEC_Tipologia profilo").HasValue == False or ParaType(el,"TEC_Tipologia profilo",doc) == "" or ParaType(el,"TEC_Tipologia profilo",doc) == None:
			param_inst_compilare.append(Para(type_el,"TEC_Tipologia profilo"))
			inst_compilare.append(type_el)
		elif ParaType(el,"TEC_Tipologia profilo",doc) == "•••COMPILARE•••":
			para_typ_check.add("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(type_el.Id),Para(type_el,"TEC_Tipologia profilo").Definition.Name))
		elif ParaType(el,"TEC_Tipologia profilo",doc) == "ND":
			para_ND_trasforma.append(Para(type_el,"TEC_Tipologia profilo"))
	else:
		try:
			para_eccesso.append(Para(type_el,"TEC_Tipologia profilo"))
		except:
			pass

#-----TEC_Materiale (txt_type)

	if opera_el in ["CV","IM","MA","MO","PV","SI"] and parteopera_el in ["AN","AC","FP","FO","FS","IC","IA","PD","PI","SP","IE","BA","ES","BS","BC"] and elemento_el in ["STL","SEL","TRV","CAP","OPO","PEN","SAR","TIM","DIA","MPL","PAL","PZF","CEN","POZ","BIN","CAS","COR","CUN","MUS","SGE","SOL","SSB","TRA","VEL","CNT","LOR","PUN","SBL","TAN","BAG","PUL","RIS","MAN","MDA","MFR","PAR","BLI","CNP","CAE","CAV","DIE","IMT","PLI","PZE","BAN","BPT","MON","TEC","TCM","LMC","PAN","RAN","TTA","NJE"]:

		if Para(type_el,"TEC_Materiale").HasValue == False or ParaType(el,"TEC_Materiale",doc) == "" or ParaType(el,"TEC_Materiale",doc) == None:
			param_inst_compilare.append(Para(type_el,"TEC_Materiale"))
			inst_compilare.append(type_el)
		elif ParaType(el,"TEC_Materiale",doc) == "•••COMPILARE•••":
			para_typ_check.add("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(type_el.Id),Para(type_el,"TEC_Materiale").Definition.Name))
		elif ParaType(el,"TEC_Materiale",doc) == "ND":
			para_ND_trasforma.append(Para(type_el,"TEC_Materiale"))
	else:
		try:
			para_eccesso.append(Para(el,"TEC_Materiale"))
		except:
			pass



#-----TEC_Utilizzo (txt_inst)

	if opera_el in ["IM"] and parteopera_el in ["IE"] and elemento_el in ["BLI","CAB","CNP","CAE","CAV","CEE","DIE","IMT","MUL","PZE","QEB","QEM"]:

		if Para(el,"TEC_Utilizzo").HasValue == False or ParaInst(el,"TEC_Utilizzo") == "" or ParaInst(el,"TEC_Utilizzo") == None:
			param_inst_compilare.append(Para(el,"TEC_Utilizzo"))
			inst_compilare.append(el)
		elif ParaInst(el,"TEC_Utilizzo") == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),Para(el,"TEC_Utilizzo").Definition.Name))
		elif ParaInst(el,"TEC_Utilizzo") == "ND":
			para_ND_trasforma.append(Para(el,"TEC_Utilizzo"))
	else:
		try:
			para_eccesso.append(Para(el,"TEC_Utilizzo"))
		except:
			pass

#-----TEC_Classificazione duso stradale (txt_inst)

	if opera_el in ["CV","MA","PV"] and parteopera_el in ["IC","IA","BA"] and elemento_el in ["BIN","COR","CUN","MUS","PPZ"]:

		if el.LookupParameter(parastrano.Name).HasValue == False or el.LookupParameter(parastrano.Name).AsString() == "" or el.LookupParameter(parastrano.Name).AsString() == None:
			param_inst_compilare.append(el.LookupParameter(parastrano.Name))
			inst_compilare.append(el)
		elif el.LookupParameter(parastrano.Name).AsString() == "•••COMPILARE•••":
			para_inst_check.append("{} - {} - {} - {}_ {} --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id),el.LookupParameter(parastrano.Name).Definition.Name))
		elif el.LookupParameter(parastrano.Name).AsString() == "ND":
			para_ND_trasforma.append(el.LookupParameter(parastrano.Name))
	else:
		try:
			para_eccesso.append(el.LookupParameter(parastrano.Name))
		except:
			pass

num_geo_m = ConvUnitsFM(999)
num_geo_mq = ConvUnitsFqMq(999)


if value_form == "P1 - Compilazione Di Default":
	t_Compilare = Transaction(doc,"Inserimento •••COMPILARE•••") # type: ignore
	t_Compilare.Start()
	testox = "•••COMPILARE•••"
	for e,p in zip(inst_compilare,param_inst_compilare):
		p.Set(testox)


	for np in param_inst_999:
		if "Area" in np.Definition.Name or "Dimensione Maglia" in np.Definition.Name:
			np.Set(num_geo_mq)
		elif "Pendenza"in np.Definition.Name:
			np.Set(57.289961630754)
		elif "Peso"in np.Definition.Name:
			np.Set(3277559.05511811)
		elif "Volume"in np.Definition.Name:
			np.Set(35279.3520547671)
		elif "Incidenza"in np.Definition.Name:
			np.Set(28.288529745408)
			
		else:
			np.Set(num_geo_m)

	t_Compilare.Commit()

list_para_ele = []


para_eccesso_clean = []
para_num_eccesso_clean = []

for para in para_eccesso:
	if para != None:
		para_eccesso_clean.append(para)

for para in para_num_eccesso:
	if para != None and para.HasValue:
		para_num_eccesso_clean.append(para)


if value_form == "P2 - Verifica Compilare - 999":

	t_Rimuovere = Transaction(doc,"Rimuovere Eccesso") # type: ignore
	t_Rimuovere.Start()

	for para in para_eccesso_clean:
		para.Set("")
	t_Rimuovere.Commit()

	if len(para_inst_check) != 0:
		output.print_md(	'# - COMPILARE INSTANZA RIMANENTI :red_circle:')

		for n in para_inst_check:
			output.print_md(	'## {} '.format(n))
	else:
		output.print_md(	'# - NESSUN COMPILARE INSTANZA RIMANENTE :white_heavy_check_mark:')
		
	if len(para_typ_check)!= 0:
		output.print_md(	'# - COMPILARE TIPI RIMANENTI :red_circle:')
		for n in para_typ_check:
			output.print_md(	'## {} '.format(n))
	else:
		output.print_md(	'# - NESSUN COMPILARE TIPO RIMANENTE :white_heavy_check_mark:')
	

	if len(para_num_eccesso_clean)!=0:
		output.print_md(	'# - PARAMETRI NUMERO IN ECCESSO :red_circle:')
		for n in para_num_eccesso_clean:
			output.print_md(	"{} - {} --> :heavy_multiplication_x:".format(output.linkify(n.Element.Id),n.Definition.Name))
	else:
		output.print_md(	'# - NESSUN PARAMETRO NUMERO IN ECCESSO :white_heavy_check_mark:')



if value_form == "P3 - Trasforma ND e 111":

	t_Transforma = Transaction(doc,"Transforma") # type: ignore
	t_Transforma.Start()
	
	if len(para_ND_trasforma)!=0:
		for paratr in para_ND_trasforma:
			paratr.Set("-")
	if len(para_111_trasforma)!=0:
		for paranu in para_111_trasforma:
			paranu.Set(0)
	t_Transforma.Commit()

"""Model Check Modelli ANAS L-4"""

__title__= 'Model Checker\nAnas L-4'
__author__= 'Luca Rosati'

import clr

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *

clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI.Selection import *

clr.AddReference('RevitNodes')
import Revit
clr.ImportExtensions(Revit.GeometryConversion)
clr.ImportExtensions(Revit.Elements)

clr.AddReference('RevitServices')
from System.Collections.Generic import *

from pyrevit import forms
from pyrevit import script

doc =__revit__.ActiveUIDocument.Document
uidoc =__revit__.ActiveUIDocument

rapp = doc.Application

#-------------------------------------DEFINIZIONI

def ExitScript(check):
	if not check:
		script.exit()

def ConvUnitsFM(number): #Feet to m
	output = number/0.3048
	return output

def ParaBuilt (element,builtinparameter):
	parameter = element.get_Parameter(BuiltInParameter.builtinparameter)
	return parameter

def Para(element,paraname):
	parameter = element.LookupParameter(paraname)
	return parameter

def ParaInst(element,paraname):
	if element.LookupParameter(paraname).StorageType == StorageType.Double:
		value = element.LookupParameter(paraname).AsDouble()
	elif element.LookupParameter(paraname).StorageType == StorageType.ElementId:
		value = element.LookupParameter(paraname).AsElementId()
	elif element.LookupParameter(paraname).StorageType == StorageType.String:
		value = element.LookupParameter(paraname).AsString()
	elif element.LookupParameter(paraname).StorageType == StorageType.Integer:
		value = element.LookupParameter(paraname).AsInteger()
	elif element.LookupParameter(paraname).StorageType == None:
		value = "Da Compilare"
	return value

def ParaType(element,paraname,document):
	element_type = document.GetElement(element.GetTypeId())
	if element_type.LookupParameter(paraname).StorageType == StorageType.Double:
		value = element_type.LookupParameter(paraname).AsDouble()
	elif element_type.LookupParameter(paraname).StorageType == StorageType.ElementId:
		value = element_type.LookupParameter(paraname).AsElementId()
	elif element_type.LookupParameter(paraname).StorageType == StorageType.String:
		value = element_type.LookupParameter(paraname).AsString()
	elif element_type.LookupParameter(paraname).StorageType == StorageType.Integer:
		value = element_type.LookupParameter(paraname).AsInteger()
	elif element_type.LookupParameter(paraname).StorageType == None:
		value = "Da Compilare"
	return value

#-------------------------------------FINESTRA OUTPUT

output = script.get_output()

output.resize(1200,800)


#-------------------------------------INFORMAZIONI MODELLO
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


#-------------------------------------INFORMAZIONI DI PROGETTO

infoparameters = ["IDP_Codice Intervento","SIT_Sistema di Coordinate","IDP_Nome opera","IDP_Struttura territoriale","SIT_Codice Strada","SIT_Comune","SIT_Regione"]

if Para(info,"IDP_Codice Intervento").HasValue == False or ParaInst(info,"IDP_Codice Intervento") == "":
	info_cin = "IDP_Codice Intervento - Da Compilare"
else:
	info_cin = "IDP_Codice Intervento = {}".format(ParaInst(info,"IDP_Codice Intervento"))


if Para(info,"SIT_Sistema di Coordinate").HasValue == False or ParaInst(info,"SIT_Sistema di Coordinate") == "":
	info_sdc = "SIT_Sistema di Coordinate - Da Compilare"
else:
	info_sdc = "SIT_Sistema di Coordinate = {}".format(ParaInst(info,"SIT_Sistema di Coordinate"))


if Para(info,"IDP_Nome opera").HasValue == False or ParaInst(info,"IDP_Nome opera") == "":
	info_nop = "IDP_Nome opera - Da Compilare"
else:
	info_nop = "IDP_Nome opera = {}".format(ParaInst(info,"IDP_Nome opera"))


if Para(info,"IDP_Struttura territoriale").HasValue == False or ParaInst(info,"IDP_Struttura territoriale") == "":
	info_ste = "IDP_Struttura territoriale - Da Compilare"
else:
	info_ste = "IDP_Struttura territoriale = {}".format(ParaInst(info,"IDP_Struttura territoriale"))


if Para(info,"SIT_Codice Strada").HasValue == False or ParaInst(info,"SIT_Codice Strada") == "":
	info_cos = "SIT_Codice Strada - Da Compilare"
else:
	info_cos = "SIT_Codice Strada = {}".format(ParaInst(info,"SIT_Codice Strada"))


if Para(info,"SIT_Comune").HasValue == False or ParaInst(info,"SIT_Comune") == "":
	info_com = "SIT_Comune - Da Compilare"
else:
	info_com = "SIT_Comune = {}".format(ParaInst(info,"SIT_Comune"))


if Para(info,"SIT_Regione").HasValue == False or ParaInst(info,"SIT_Regione") == "":
	info_reg = "SIT_Regione - Da Compilare"
else:
	info_reg = "SIT_Regione = {}".format(ParaInst(info,"SIT_Regione"))


proj_info =[info_cin,info_sdc,info_nop,info_ste,info_cos,info_com,info_reg]


#WARNING_00-----------------INFORMAZIONI DI PROGETTO
proj_info_errato = []

for inf in proj_info:
	if "Da Compilare" in inf:
		proj_info_errato.append(inf)

if len(proj_info_errato) != 0:
	forms.alert('WARNING 00_INFORMAZIONI DI PROGETTO\n\nI seguenti parametri in informazioni di progetto risultano vuoti e devono essere compilati\n\n{}'.format(proj_info_errato), exitscript=True)
else:
	proj_info_result = 'WARNING 01_INFORMAZIONI DI PROGETTO --> :white_heavy_check_mark:'


#WARNING_02-----------------CODICE INTERVENTO

info_cin_true = ParaInst(info,"IDP_Codice Intervento")

if info_cin_true == 'AQ 62-22 Lotto 4':
	codint_result = "WARNING 02_CODICE INTERVENTO --> :white_heavy_check_mark:"
else:
	forms.alert('WARNING 02_CODICE INTERVENTO\n\nIl parametro IDP_Codice Intervento deve essere compilato con il valore AQ 62-22 Lotto 4\n\nIDP_Codice Intervento 0 {}'.format(info_cin_true), exitscript=True)

bpointlist = FilteredElementCollector(doc).OfClass(BasePoint).WhereElementIsNotElementType().ToElements()

bpoint=bpointlist[1]

site = doc.SiteLocation
psitelist = FilteredElementCollector(doc).OfClass(ProjectLocation).WhereElementIsNotElementType().ToElements()
psite = psitelist[0]
site_name = 'Nome Posizione = {}'.format(psite.Name)
site_lat = "Latitudine = {}".format(round(site.Latitude*(180/3.141592653589793),13))
site_long = "Longitudine = {}".format(round(site.Longitude*(180/3.141592653589793),13))

bpoint_NS = "Base Point N/S = {}".format(bpoint.LookupParameter("N/S").AsValueString())
bpoint_EO = "Base Point E/O = {}".format(bpoint.LookupParameter("E/O").AsValueString())
bpoint_ele = "Base Point Elevation = {}".format(bpoint.LookupParameter("Quota altim.").AsValueString())
bpoint_angle = "Angolo con il True North = {} gradi".format(round(bpoint.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM).AsDouble()*(180/3.141592653589793),3))


#WARNING_03-----------------NOME POSIZIONE E PARAMETRO SISTEMA COORDINATE

info_sdc_true = ParaInst(info,"SIT_Sistema di Coordinate")

if len(info_sdc_true) == 9 and "EPSG" in info_sdc_true and info_sdc_true in site_name:
	sitename_result = "WARNING 03_SISTEMA COORDINATE --> :white_heavy_check_mark:"
else:
	forms.alert('WARNING 03_SISTEMA COORDINATE\n\nIl Codice inserito nel parametro SIT_Sistema di Coordinate non corretto o non coincide con il nome della posizione condivisa o viceversa\n\nSIT_Sistema di Coordinate = {}\n Nome Posizione = {}'.format(info_sdc,site_name), exitscript=True)

site_info = [site_name,site_lat,site_long,bpoint_NS,bpoint_EO,bpoint_ele,bpoint_angle]


#-------------------------------------FASI

phases_true_name = ["Esistente","Nuova costruzione"]

d_phases = doc.Phases
phases_name = []

for p in d_phases:
	phases_name.append(p.Name)

result_ph = []
ph_notcompl = []

#WARNING_04-----------------FASI

for pn in phases_name:
	if pn not in phases_true_name:
		forms.alert('WARNING 04_FASI\n\nLe seguenti fasi non sono state impostate correttamente o hanno un nome diverso\n\n {}'.format(pn), exitscript=True)
	else:
		result_ph = "WARNING 04_FASI --> :white_heavy_check_mark:"
		break


#-------------------------------------WORKSET

work_true_name = ["00_RIL","01_Griglie e livelli","02_Temporaneo","03_Nascosto","04_STR","05_TRA","06_GET","07_AMB","08_IMP"]

d_wset = FilteredWorksetCollector(doc)

wset_name = []

for c in d_wset:
	wset_k = c.Kind
	if wset_k == WorksetKind.UserWorkset:
		wset_name.append(c.Name)

result_ws = []

#WARNING_05-----------------WORKSET

for wn in wset_name:
	if wn not in work_true_name:
		forms.alert('WARNING 05_WORKSET\n\nI seguenti Wokset non sono Impostati correttamente o hanno un nome sbagliato\n\n {}'.format(wn), exitscript=True)
	else:
		result_ws = "WARNING 05_WORKSET --> :white_heavy_check_mark:"
		break


#-------------------------------------EXP.VIEWS

pvp = ParameterValueProvider(ElementId(634609))
fng = FilterStringEquals()
ruleValue = 'D_ESPORTAZIONI_IFC'
fRule = FilterStringRule(pvp,fng,ruleValue,True)

filter = ElementParameterFilter(fRule)

exp_views_coll = FilteredElementCollector(doc).OfClass(View3D).WherePasses(filter).WhereElementIsNotElementType().ToElements()

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


#WARNING_06-----------------NOME E PRESENZA VISTE ESPORTAZIONE

if len(exp_view_check) == len(exp_views):
	exp_view_result = "WARNING 06_VISTE ESPORTAZIONE IFC --> :white_heavy_check_mark:"
else:
	forms.alert('WARNING 06_VISTE ESPORTAZIONE IFC\n\nLe seguenti viste di esportazione non sono correttemente impostate o rinominate\n\n {}'.format(exp_view_false), exitscript=True)

#-------------------------------------ELEMENTI

for ev in exp_views_ele:
	if "FED" in ev.Name:
		ex_view_fg = ev

#WARNING_07-----------------VISTA ESPORTAZIONE FEDERATO

if ex_view_fg.get_Parameter(BuiltInParameter.VIEW_PHASE).AsValueString()=="Nuova costruzione" and ex_view_fg.get_Parameter(BuiltInParameter.VIEW_PHASE_FILTER).AsValueString()=="Mostra completo":
	exp_viewed_result = "WARNING 07_VISTA EXP FEDERATO --> :white_heavy_check_mark:"
else:
	forms.alert('WARNING 07_VISTA EXP FEDERATO\n\nLa vista di esportazione FED non ha le fasi impostate correttamente\n\n Correggere la vista con le seguenti impostazioni\n\nFiltro delle fasi = Mostra completo\nFase = Nuova costruzione', exitscript=True)

doc_el = FilteredElementCollector(doc,ex_view_fg.Id).WhereElementIsNotElementType().ToElements()


cat = []
cnamelist = []
catname = []

clean_el = []

for el in doc_el:
	try:
		if el.Category.CategoryType == CategoryType.Model and "dwg"  not in el.Category.Name and el.Category.SubCategories.Size > 0 and el.Category.CanAddSubcategory:
			clean_el.append(el)
	except:
		pass

if len(clean_el) == 0:
	forms.alert('WARNING 07_VISTA EXP FEDERATO\n\nLa vista di esportazione FED non contiene nessun elemento\n\n Correggere  le impostazioni della vista', exitscript=True)

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
		cat_result = "WARNING 08_CATEGORIE MODELLO DATI --> :white_heavy_check_mark:"
	else:
		cat_false.append(c)

#WARNING_08-----------------CATEGORIA

if len(cat_false) != 0:
	forms.alert('WARNING 08_CATEGORIE\n\nLe seguenti categorie non sono presenti nel modello dati,\n\n{}\n\nEliminare gli elementi prima di procedere'.format(cat_false), exitscript=True)

#WARNING_09-----------------ELEMENTI RILIEVO

worksetril_errata = []

doc_el_tot = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

work_ele = []

for elt in doc_el_tot:
	try:
		if elt.Category.CategoryType == CategoryType.Model and "dwg" not in elt.Category.Name and elt.Category.SubCategories.Size > 0 and elt.Category.CanAddSubcategory == True and "Detail" not in elt.Category.Name and "Line" not in elt.Category.Name:
			work_ele.append(elt)
	except:
		pass

numfe = 0

for we in work_ele:
	try:
		type_we = doc.GetElement(we.GetTypeId())
		category_we = we.Category.Name
		opera_we = type_we.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsValueString()
		parteopera_we = type_we.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString()
		elemento_we = type_we.get_Parameter(BuiltInParameter.ALL_MODEL_DESCRIPTION).AsValueString()
		we_id = we.Id
		try:
			symbol = we.Symbol
			type_we_name = type_we.FamilyName
		except:
			type_we_name = we.Name
		
		if we.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString() == "00_RIL"and we.Category.Name != "Nuvole di punti" and "iferiment" not in we.Category.Name:
			numfe += 1
			worksetril_errata.append(":heavy_multiplication_x: {} - {} - {} - {} - {} - {} - {}".format(numfe,category_we,type_we_name,opera_we,parteopera_we,elemento_we,output.linkify(we_id)))

		if we.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString() == "01_Griglie e livelli" and we.Category.Name != "Nuvole di punti" and "iferiment" not in we.Category.Name:
				numfe += 1
				worksetril_errata.append(":heavy_multiplication_x: {} - {} - {} - {} - {} - {} - {}".format(numfe,category_we,type_we_name,opera_we,parteopera_we,elemento_we,output.linkify(we_id)))
	except:
		pass

if len(worksetril_errata) != 0:
	output.print_md(	'# :red_circle: WARNING 09_ASSOCIAZIONE WORKSET RILIEVO/GRIGLIE E LIVELLI')
	output.print_md(	'##I seguenti Elementi hanno Associato erronemanete il workset Rilievo/Griglie e Livelli')
	for i in worksetril_errata:
		output.print_md(	'###{}'.format(i))
else:
	result_awril = "WARNING 09_ASSOCIAZIONE WORKSET RILIEVO/GRIGLIE E LIVELLI --> :white_heavy_check_mark:"

if len(worksetril_errata) != 0:
	script.exit()

#-------------------------------------LIVELLI


levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()

level_ns =[]

for l in levels:
	level_ns.append(l.Name)

level_true = ["LIV_RIL","LIV_TRA","LIV_STR","LIV_SLM"]

level_check = []
level_check_false = []

for l in level_ns:
    presente = False
    for i in level_true:
        if i in l:
            presente = True
            break
    if presente:
        level_check.append("{} --> V".format(l))
    else:
        level_check_false.append(l)



#VERIFICA_00-----------------NOME LIVELLI

if len(level_check_false) == 0:
	result_level = "VERIFICA 00_NOME LIVELLI --> :white_heavy_check_mark:"
else:
	forms.alert('VERIFICA 00_NOME LIVELLI\n\nI seguenti livelli risultano non correttamente rinominati\n\n {}'.format(level_check_false), exitscript=True)


#-------------------------------------ASSOCIAZIONE FASE

numfe = 0
assfase_errata = []
assLivello_errata = []

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
		assfase_errata.append(":heavy_multiplication_x: {} - {} - {} - {} - {} - {} - {}".format(numfe,category_el,type_el_name,opera_el,parteopera_el,elemento_el,output.linkify(el_id)))

	if opera_el in ['PV','MA','SI'] and el.get_Parameter(BuiltInParameter.PHASE_CREATED).AsValueString() != "Esistente":
		numfe += 1
		assfase_errata.append(":heavy_multiplication_x: {} - {} - {} - {} - {} - {} - {}".format(numfe,category_el,type_el_name,opera_el,parteopera_el,elemento_el,output.linkify(el_id)))
	
	if opera_el in ["MO"] and parteopera_el in ["BE","CP","EA","ES","MN","TR"] and elemento_el in ["BRE","CLI","TSS","IFS","TEC","TCM","ACC","SCA","VCM","IDO","SEM","SUM","TIG"] and el.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM).AsValueString() != "LIV_SLM_000":
		numfe += 1
		assLivello_errata.append(":heavy_multiplication_x: {} - {} - {} - {} - {} - {} - {}".format(numfe,category_el,type_el_name,opera_el,parteopera_el,elemento_el,output.linkify(el_id)))


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


#-------------------------------------HOST

	try:
		if el.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_PARAM).AsValueString() != "<non associato>":
			host_check = True
		else:
			num_h += 1
			host_errato.append("{} - {} - {} - {} - {} - {} - {}".format(num_h,category_el,type_el_name,opera_el,parteopera_el,elemento_el,output.linkify(el_id)))
	except:
		pass


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


#VERIFICA_05-----------------ASSOCIAZIONE HOST

if len(host_errato) != 0:
	output.print_md(	'# :red_circle: VERIFICA 05_ASSOCIAZIONE HOST')
	output.print_md(	'## I seguenti Elementi hanno il parametro Host non associato')
	for h in host_errato:
		output.print_md(	'###{}'.format(h))

else:
	result_host = "VERIFICA 05_ASSOCIAZIONE HOST --> :white_heavy_check_mark:"

if len(host_errato) != 0:
	script.exit()


#VERIFICA_06-----------------ASSOCIAZIONE LIVELLO

if len(assLivello_errata) != 0:
	output.print_md(	'# :red_circle: VERIFICA 06_ASSOCIAZIONE LIVELLO SENSORI')
	output.print_md(	'## I seguenti Elementi hanno il livello errato')
	for h in assLivello_errata:
		output.print_md(	'###{}'.format(h))

else:
	result_aliv = "VERIFICA 06_ASSOCIAZIONE LIVELLO SENSORI --> :white_heavy_check_mark:"

if len(assLivello_errata) != 0:
	script.exit()

#-------------------------------------CLUSTER INFORMATIVI

codiceopera_errato = []
codiceWBS_errato = []
gruppoanagrafica_errato = []
id_elemento_errato = []
lor_errato = []
codiceassieme_errato = []
campatadiappartenenza_errato = []
impalcatodiappartenenza_errato = []
numstrutturacampata_errato = []
codicebms_errato = []
progettista_errato = []
area_errato = []
volume_errato = []
qsensore_errato = []
tposizione_errato = []
codicesensore_errato = []
numeroseriale_errato = []
carreggiata_errato = []
direzione_errato = []


list_clusterID_OG = [codiceopera_errato,codiceWBS_errato,gruppoanagrafica_errato,lor_errato,codiceassieme_errato,codicesensore_errato]

list_clusterIN_D = [campatadiappartenenza_errato,impalcatodiappartenenza_errato,numstrutturacampata_errato,codicebms_errato]

list_clusterGEO = [area_errato,volume_errato,qsensore_errato]

list_clusterANA = [progettista_errato]

list_clusterLOC = [carreggiata_errato,direzione_errato]

list_clusterTEC = [tposizione_errato,numeroseriale_errato]

listone = [codiceopera_errato,codiceWBS_errato,gruppoanagrafica_errato,lor_errato,codiceassieme_errato,codicesensore_errato,campatadiappartenenza_errato,impalcatodiappartenenza_errato,numstrutturacampata_errato,codicebms_errato,carreggiata_errato,direzione_errato,area_errato,volume_errato,qsensore_errato,progettista_errato,tposizione_errato,numeroseriale_errato]

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

#-------------------------------------IDENTIFICATIVO OGGETTO

	if Para(el,"IDE_Codice opera").HasValue == False or ParaInst(el,"IDE_Codice opera") == "" or len(ParaInst(el,"IDE_Codice opera"))!= 11:
		codiceopera_errato.append("{} - {} - {} - {}_ IDE_Codice opera --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))

	if Para(el,"IDE_Codice WBS").HasValue == False or ParaInst(el,"IDE_Codice WBS") == "" or len(ParaInst(el,"IDE_Codice WBS"))!= 21 and opera_el in ParaInst(el,"IDE_Codice WBS") and parteopera_el in ParaInst(el,"IDE_Codice WBS") and elemento_el in ParaInst(el,"IDE_Codice WBS"):
		codiceWBS_errato.append("{} - {} - {} - {}_ IDE_Codice WBS --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))

	if Para(el,"IDE_Gruppo anagrafica").HasValue == False or ParaInst(el,"IDE_Gruppo anagrafica") == "" or len(ParaInst(el,"IDE_Gruppo anagrafica"))!= 20 or opera_el not in ParaInst(el,"IDE_Gruppo anagrafica") or parteopera_el not in ParaInst(el,"IDE_Gruppo anagrafica") or elemento_el not in ParaInst(el,"IDE_Gruppo anagrafica"):
		gruppoanagrafica_errato.append("{} - {} - {} - {}_ IDE_Gruppo anagrafica --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))

	if Para(el,"IDE_LOR").HasValue == False or ParaInst(el,"IDE_LOR") == "" or ParaInst(el,"IDE_LOR") not in ["MOLTO BASSO","BASSO","MEDIO","ALTO"]:
		lor_errato.append("{} - {} - {} - {}_ IDE_LOR --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))

	if Para(type_el,"Codice assieme").HasValue == False or ParaType(el,"Codice assieme",doc) == "" or "Co" not in ParaType(el,"Codice assieme",doc) or "En" not in ParaType(el,"Codice assieme",doc) or "EF" not in ParaType(el,"Codice assieme",doc):
		codiceassieme_errato.append("{} - {} - {} - {}_ Codice assieme --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))

	if elemento_el in ["IFS","TEC","TCM","ACC","SCA","VCM","TSS","SEM","TIG","SUM","CLI","BRE","IDO"]:
		if Para(el,"IDE_Codice sensore").HasValue == False or ParaInst(el,"IDE_Codice sensore") == "":
			codicesensore_errato.append("{} - {} - {} - {}_ IDE_Codice sensore --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))


#-------------------------------------INFORMAZIONI 6D


	if 'XXX' not in type_el_name and category_el in ["Appoggi","Collegamenti strutturali","Modelli generici","Pavimenti","Pilastri strutturali","Telaio strutturale","Tetti"]:
		if Para(el,"INF_Campata di appartenenza").HasValue == False or ParaInst(el,"INF_Campata di appartenenza") == "" or len(ParaInst(el,"INF_Campata di appartenenza"))!= 3:
			campatadiappartenenza_errato.append("{} - {} - {} - {}_ INF_Campata di appartenenza --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))

	if 'XXX' not in type_el_name and category_el in ["Appoggi","Collegamenti strutturali","Modelli generici","Pavimenti","Pilastri strutturali","Telaio strutturale","Tetti"]:
		if Para(el,"INF_Impalcato di appartenenza").HasValue == False or ParaInst(el,"INF_Impalcato di appartenenza") == "" or len(ParaInst(el,"INF_Impalcato di appartenenza"))!= 3:
			impalcatodiappartenenza_errato.append("{} - {} - {} - {}_ INF_Impalcato di appartenenza --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))

	if 'XXX' not in type_el_name and category_el in ["Appoggi","Collegamenti strutturali","Modelli generici","Pavimenti","Pilastri strutturali","Telaio strutturale","Tetti"]:
		if Para(el,"INF_Numerazione struttura campata").HasValue == False or ParaInst(el,"INF_Numerazione struttura campata") == "" or len(ParaInst(el,"INF_Numerazione struttura campata"))!= 3:
			numstrutturacampata_errato.append("{} - {} - {} - {}_ INF_Numerazione struttura campata --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))

	if 'XXX' not in type_el_name and category_el in ["Appoggi","Collegamenti strutturali","Modelli generici","Pavimenti","Pilastri strutturali","Telaio strutturale","Tetti"]:
		if Para(el,"INF_Codice BMS").HasValue == False or ParaInst(el,"INF_Codice BMS") == "" or len(ParaInst(el,"INF_Codice BMS"))!= 23 or ParaInst(el,"INF_Campata di appartenenza") not in ParaInst(el,"INF_Codice BMS") or ParaInst(el,"INF_Impalcato di appartenenza") not in ParaInst(el,"INF_Codice BMS") or ParaInst(el,"INF_Numerazione struttura campata") not in ParaInst(el,"INF_Codice BMS") :
			codicebms_errato.append("{} - {} - {} - {}_ INF_Codice BMS --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))


#-------------------------------------ANAGRAFICA DI BASE

	if elemento_el in ['ALI','BLI','CAB','CNP','CAE','CAV','COL','COM','CEE','DIE','IMT','INS','INM','MUL','PLI','PZE','QEB','QEM','REL','REP','RIF','SCS','SEM','SDE','TRS'] and opera_el == "IM" and parteopera_el == "IE":
			if Para(el,"ANA_Progettista").HasValue == False or ParaInst(el,"ANA_Progettista") == "" or ParaInst(el,"ANA_Progettista")!= "RINA Consulting SpA":
				progettista_errato.append("{} - {} - {} - {}_ ANA_Progettista --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))


#-------------------------------------LOCALIZZAZIONE

	if elemento_el in ['AAP','ALI','AMS','BAG','BAN','BIN','BLI','BPT','CAB','CAE','CAP','CAV','CEE','CEN','CNP','CNT','COL','COM','COR','CUN','DIA','DIE','GIU','IMT','INM','ISA','LMC','LOR','MAN','MDA','MFR','MON','MPL','MUL','MUS','NJE','OPO','PAL','PAN','PAR','PEN','PLI','POZ','PPZ','PUL','PUN','PZE','PZF','QEB','QEM','RAN','REL','REP','RIF','RIS','SAR','SBL','SCS','SDE','SEL','SGE','SOL','SSB','STL','TAN','TIM','TRA','TRS','TRV','TTA','UNI','VEL']:
			if Para(el,"LOC_Carreggiata").HasValue == False or ParaInst(el,"LOC_Carreggiata") == "":
				carreggiata_errato.append("{} - {} - {} - {}_ LOC_Carreggiata --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))

	if elemento_el in ['AAP','ALI','AMS','BAG','BAN','BIN','BLI','BPT','CAB','CAE','CAP','CAV','CEE','CEN','CNP','CNT','COL','COM','COR','CUN','DIA','DIE','GIU','IMT','INM','ISA','LMC','LOR','MAN','MDA','MFR','MON','MPL','MUL','MUS','NJE','OPO','PAL','PAN','PAR','PEN','PLI','POZ','PPZ','PUL','PUN','PZE','PZF','QEB','QEM','RAN','REL','REP','RIF','RIS','SAR','SBL','SCS','SDE','SEL','SGE','SOL','SSB','STL','TAN','TIM','TRA','TRS','TRV','TTA','UNI','VEL']:
			if Para(el,"LOC_Direzione").HasValue == False or ParaInst(el,"LOC_Direzione") == "":
				direzione_errato.append("{} - {} - {} - {}_ LOC_Direzione --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))


#-------------------------------------GEOMETRICO

	if elemento_el in ["CEN","CEE","PZF","PLI","POZ","QEB","QEM","RAN","SOL","SSB"]:
		if Para(el,"GEO_Area").HasValue == False or ParaInst(el,"GEO_Area") == 0:
			area_errato.append("{} - {} - {} - {}_ GEO_Area --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))

	if elemento_el in ['ALI','BAG','BLI','CAB','CNP','CAS','CAV','COL','COM','COR','DIA','DIE','IMT','INS','INM','MPL','MAN','MDA','MFR','PPZ','PAL','PAR','PZF','PLI','POZ','PUL','REL','REP','RIF','SCS','SDE','SGE','SOL','SSB','SAR','SEL','TRS','TRV','TRA']:
		if Para(el,"GEO_Volume").HasValue == False or ParaInst(el,"GEO_Volume") == 0:
			volume_errato.append("{} - {} - {} - {}_ GEO_Volume --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))

	if elemento_el in ['ACC','IFS','SCA','SEM','TEC','TCM','VCM','TSS','TIG','SUM','CLI','BRE','IDO']:
		if Para(el,"GEO_Quota sensore").HasValue == False or ParaInst(el,"GEO_Quota sensore") == 0:
			qsensore_errato.append("{} - {} - {} - {}_ GEO_Quota sensore --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))

#-------------------------------------TECNICO

	if elemento_el in ['ACC','IFS','SCA','SEM','TEC','TCM','VCM','TIG','SUM']:
		if Para(el,"TEC_Posizione").HasValue == False or ParaInst(el,"TEC_Posizione") == "":
			tposizione_errato.append("{} - {} - {} - {}_ TEC_Posizione --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))

	if elemento_el in ["IFS","TEC","TCM","ACC","SCA","VCM","TSS","SEM","TIG","SUM","CLI","BRE","IDO"]:
		if Para(el,"TEC_Numero seriale").HasValue == False or ParaInst(el,"TEC_Numero seriale") != "-":
			numeroseriale_errato.append("{} - {} - {} - {}_ TEC_Numero seriale --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))

#CHECK_01-----------------IDENTIFICATIVO OGGETTO

for l in list_clusterID_OG:
	if len(l) != 0:
		output.print_md(	'# :keycap_1: CHECK 01_CLUSTER IDENTIFICATIVO OGGETTO - Errore nella valorizzazione dei parametri:')
		break

for l in list_clusterID_OG:
	if len(l) != 0:
		for i in l:
			output.print_md(	'###{}'.format(i))


#CHECK_02-----------------INFORMAZIONI 6D


for l in list_clusterIN_D:
	if len(l) != 0:
		output.print_md(	'# :keycap_2: CHECK 02_CLUSTER INFORMAZIONI 6D - Errore nella Valorizzazione dei parametri:')
		break


for l in list_clusterIN_D:
	if len(l) != 0:
		for i in l:
			output.print_md(	'###{}'.format(i))



#CHECK_03-----------------ANAGRAFICA DI BASE

for l in list_clusterANA:
	if len(l) != 0:
		output.print_md(	'# :keycap_3: CHECK 03_CLUSTER ANAGRAFICA DI BASE - Errore nella Valorizzazione dei parametri:')
		break


for l in list_clusterANA:
	if len(l) != 0:
		for i in l:
			output.print_md(	'###{}'.format(i))



#CHECK_04-----------------LOCALIZZAZIONE

for l in list_clusterLOC:
	if len(l) != 0:
		output.print_md(	'# :keycap_3: CHECK 04_CLUSTER LOCALIZZAZIONE - Errore nella Valorizzazione dei parametri:')
		break


for l in list_clusterLOC:
	if len(l) != 0:
		for i in l:
			output.print_md(	'###{}'.format(i))



#CHECK_05-----------------GEOMETRICO

for l in list_clusterGEO:
	if len(l) != 0:
		output.print_md(	'# :keycap_4: CHECK 05_CLUSTER GEOMETRICO - Errore nella Valorizzazione dei parametri:')
		break

for l in list_clusterGEO:
	if len(l) != 0:
		for i in l:
			output.print_md(	'###{}'.format(i))


#CHECK_06-----------------TECNICO

for l in list_clusterTEC:
	if len(l) != 0:
		output.print_md(	'# :keycap_5: CHECK 06_CLUSTER TECNICO - Errore nella Valorizzazione dei parametri:')
		break

for l in list_clusterTEC:
	if len(l) != 0:
		for i in l:
			output.print_md(	'###{}'.format(i))


for lista in listone:
	if len(lista) != 0:
		script.exit()

#-------------------------------------ID ELEMENTO

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

	try:
		if Para(el,"IfcGUID").HasValue == False or ParaInst(el,"IfcGUID") == "" or len(ParaInst(el,"IfcGUID"))!= 22:
			id_elemento_errato.append("{} - {} - {} - {}_ IfcGUID --> :heavy_multiplication_x:".format(category_el,type_el_name,opera_el,output.linkify(el_id)))
	except:
		forms.alert("Effettuare un'esportazione IFC in modo da creare per tutti gli elementi il parametro IfcGUID rieseguire poi il Check",exitscript=True)

#CHECK_07-----------------ID Elemento

if len(id_elemento_errato) != 0:
	for i in id_elemento_errato:
		output.print_md(	'###{}'.format(i))

if len(id_elemento_errato) != 0:
	script.exit()

#COMPILAZIONE-----------------Parametro Nascosti a Commenti

pvpV = ParameterValueProvider(ElementId(-1005112))
fngV = FilterStringEquals()
ruleValueV = 'NASCOSTO'
fRuleV = FilterStringRule(pvpV,fngV,ruleValueV,True)

filterV = ElementParameterFilter(fRuleV)

exp_views_collV = FilteredElementCollector(doc).OfClass(View3D).WherePasses(filterV).WhereElementIsNotElementType().ToElements()

if len(exp_views_collV)!=0:
	IDV = exp_views_collV[0]
	elemsnasc = FilteredElementCollector(doc,IDV.Id).OfClass(Floor).WhereElementIsNotElementType().ToElements()
	if len(elemsnasc) != 0:
		t = Transaction(doc,"compila commento")
		t.Start()
		try:
			for ev in elemsnasc:
				ev.LookupParameter("Commenti").Set("Nascosti")
		except:
			t.RollBack()
		else:
			t.Commit()
	else:
		pass
else:
	pass

#print(file_info,proj_info,site_info,codint_result,result_ph,result_ws,exp_view_check,exp_view_result,exp_viewed_result,n_el,cat_result,result_afase,result_class,result_nomen)

output.resize(1200,800)

warn_result = [proj_info_result,codint_result,sitename_result,result_ph,result_ws,exp_view_result,exp_viewed_result,cat_result,result_awril]

veri_result = [result_level,result_afase,result_workset,result_class,result_nomen,result_host,result_aliv]

check_result = ["CHECK 01_CLUSTER IDENTIFICATIVO OGGETTO --> :white_heavy_check_mark:","CHECK 02_CLUSTER INFORMAZIONI 6D --> :white_heavy_check_mark:","CHECK 03_CLUSTER ANAGRAFICA DI BASE --> :white_heavy_check_mark:","CHECK 04_CLUSTER LOCALIZZAZIONE --> :white_heavy_check_mark:","CHECK 05_CLUSTER GEOMETRICO --> :white_heavy_check_mark:","CHECK 06_CLUSTER TECNICO --> :white_heavy_check_mark:","CHECK 07_ID ELEMENTO --> :white_heavy_check_mark:"]

output.print_md(	'# - INFORMAZIONI MODELLO :black_square_button:')

for n in file_info:
	output.print_md(	'## {} '.format(n))

output.print_md(	'-----------------------')

output.print_md(	'# - INFORMAZIONI DI PROGETTO :bookmark_tabs:')

for p in proj_info:
	output.print_md(	'## {} '.format(p))

output.print_md(	'-----------------------')

output.print_md(	'# - SISTEMA DI COORDINATE :round_pushpin:')

for s in site_info:
	output.print_md(	'## {} '.format(s))

output.print_md(	'-----------------------')

output.print_md(	'# - WARNINGS PRINCIPALI :warning:')

for wp in warn_result:
	output.print_md(	'## {} '.format(wp))

output.print_md(	'-----------------------')

output.print_md(	'# - VERIFICHE PRELIMINARI :clipboard:')

for ve in veri_result:
	output.print_md(	'## {} '.format(ve))

output.print_md(	'-----------------------')

output.print_md(	'# - CHECK CLUSTER PARAMETRI :flashlight:')

for cr in check_result:
	output.print_md(	'## {} '.format(cr))

output.print_md(	'-----------------------')

output.print_md(	"# - PROCEDERE CON L'ESPORTAZIONE DEI FILE IFC :clinking_beer_mugs:")
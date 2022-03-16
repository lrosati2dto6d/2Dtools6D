"""This Algorithm allows, based on the indication of a Air Flow rate, to indicate the optimal sections based on a table of Pressure Drops"""

__title__= "Duct Section\nSizing"
__author__= "Luca Rosati-Davide Lavieri-Fausto Baronti"

import System
import clr

from collections import defaultdict
from pyrevit import HOST_APP
from pyrevit.framework import List
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import script

dptt = []

val_min =(0.0320,0.0340,0.0370,0.0380,0.0380,0.0390,0.0390,0.0390,0.0390,0.0400,0.0403,0.0406,0.0409,0.0412,0.0415,0.0418,0.0421,0.0424,0.0427,0.0430,0.0431,0.0432,0.0433,0.0434,0.0435,0.0436,0.0437,0.0438,0.0439,0.0440,0.0441,0.0442,0.0443,0.0444,0.0445,0.0446,0.0447,0.0448,0.0449,0.0450,0.0451,0.0452,0.0453,0.0454,0.0455,0.0456,0.0457,0.0458,0.0459,0.0460,0.0460,0.0460,0.0460,0.0460,0.0463,0.0463,0.0463,0.0463,0.0463,0.0466,0.0466,0.0466,0.0466,0.0466,0.0469,0.0469,0.0469,0.0469,0.0469,0.0472,0.0472,0.0472,0.0472,0.0472,0.0475,0.0475,0.0475,0.0475,0.0475,0.0478,0.0478,0.0478,0.0478,0.0478,0.0481,0.0481,0.0481,0.0481,0.0481,0.0484,0.0484,0.0484,0.0484,0.0484,0.0487,0.0487,0.0487,0.0487,0.0487,0.0490,0.0490,0.0490,0.0490,0.0490,0.0492,0.0492,0.0492,0.0492,0.0492,0.0494,0.0494,0.0494,0.0494,0.0494,0.0496,0.0496,0.0496,0.0496,0.0496,0.0498,0.0498,0.0498,0.0498,0.0498,0.0500,0.0500,0.0500,0.0500,0.0500,0.0502,0.0502,0.0502,0.0502,0.0502,0.0504,0.0504,0.0504,0.0504,0.0504,0.0506,0.0506,0.0506,0.0506,0.0506,0.0508,0.0508,0.0508,0.0508,0.0508,0.0510,0.0510,0.0510,0.0510,0.0510,0.0511,0.0511,0.0511,0.0511,0.0511,0.0512,0.0512,0.0512,0.0512,0.0512,0.0513,0.0513,0.0513,0.0513,0.0513,0.0514,0.0514,0.0514,0.0514,0.0514,0.0515,0.0515,0.0515,0.0515,0.0515,0.0516,0.0516,0.0516,0.0516,0.0516,0.0517,0.0517,0.0517,0.0517,0.0517,0.0518,0.0518,0.0518,0.0518,0.0518,0.0519,0.0519,0.0519,0.0519,0.0519,0.0520)

val_max = (0.0500,0.0500,0.0500,0.0510,0.0520,0.0540,0.0550,0.0560,0.0570,0.0570,0.0575,0.0580,0.0585,0.0590,0.0595,0.0600,0.0605,0.0610,0.0615,0.0620,0.0623,0.0626,0.0629,0.0632,0.0635,0.0638,0.0641,0.0644,0.0647,0.0650,0.0653,0.0656,0.0659,0.0662,0.0665,0.0668,0.0671,0.0674,0.0677,0.0680,0.0682,0.0684,0.0686,0.0688,0.0690,0.0692,0.0694,0.0696,0.0698,0.0700,0.0700,0.0700,0.0700,0.0700,0.0706,0.0706,0.0706,0.0706,0.0706,0.0712,0.0712,0.0712,0.0712,0.0712,0.0718,0.0718,0.0718,0.0718,0.0718,0.0724,0.0724,0.0724,0.0724,0.0724,0.0730,0.0730,0.0730,0.0730,0.0730,0.0736,0.0736,0.0736,0.0736,0.0736,0.0742,0.0742,0.0742,0.0742,0.0742,0.0748,0.0748,0.0748,0.0748,0.0748,0.0754,0.0754,0.0754,0.0754,0.0754,0.0760,0.0760,0.0760,0.0760,0.0760,0.0764,0.0764,0.0764,0.0764,0.0764,0.0768,0.0768,0.0768,0.0768,0.0768,0.0772,0.0772,0.0772,0.0772,0.0772,0.0776,0.0776,0.0776,0.0776,0.0776,0.0780,0.0780,0.0780,0.0780,0.0780,0.0784,0.0784,0.0784,0.0784,0.0784,0.0788,0.0788,0.0788,0.0788,0.0788,0.0792,0.0792,0.0792,0.0792,0.0792,0.0796,0.0796,0.0796,0.0796,0.0796,0.0800,0.0800,0.0800,0.0800,0.0800,0.0802,0.0802,0.0802,0.0802,0.0802,0.0804,0.0804,0.0804,0.0804,0.0804,0.0806,0.0806,0.0806,0.0806,0.0806,0.0808,0.0808,0.0808,0.0808,0.0808,0.0810,0.0810,0.0810,0.0810,0.0810,0.0812,0.0812,0.0812,0.0812,0.0812,0.0814,0.0814,0.0814,0.0814,0.0814,0.0816,0.0816,0.0816,0.0816,0.0816,0.0818,0.0818,0.0818,0.0818,0.0818,0.0820)

for min_v,max_v in zip(val_min,val_max):
	dptt.append(min_v+max_v/2)

press = range(100,20100,100)


port = forms.ask_for_one_item(
    press,
    default=press[19],
    prompt='Select Required Flow Rate (m3/h)',
    title='Duct Sizing'
)

ppmin = []

for p,d,vmi,vma in zip(press,dptt,val_min,val_max):
	if port == p:
		dpt = d
		val_mi = vmi
		val_ma = vma		
		

duct_m = "Galvanised Sheet Metal"
rug = 0.09
gas_t = "Air 20 gr C"
visc = 1.5*10**-5
dens = 1.2

listnum = range(25,4000,25)

diaeq = [] 

for n in listnum:
	r = 0.6376*(10**7)*(0.11*((rug/n)+(192.3*n*visc/port))**0.25)*dens*(port**2)/(n**5) 
	if r <= dpt:
		diaeq.append(n)	

deq = diaeq[0]

sez = round(pow(2,0.25)/1.3*deq,0)

if sez < 1000 and sez >= 100:
	num_c = sez-int(str(sez)[0])*100
elif sez > 1000:
	num_c = sez-int(str(sez)[0:2])*100
elif sez > 0 and sez < 100:
	num_c = sez-int(str(sez)[0])*1

sez_app= sez

if num_c in range(1,24,1) or num_c in range(50,74,1):
	while sez_app%50!=0:
		sez_app -= 1
else:
	while sez_app%50!=0:
		sez_app += 1
	
b = sez_app
h1 = sez_app+50
h2 = sez_app-50
h3 = sez_app+100
h4 = sez_app-100

vel = port/3600.0/(sez_app/1000)**2

def deqn(base,altezza):
	output = (1.3*(base * altezza)**0.625)/(base+altezza)**0.25
	risultato = 0.6376*(10**7)*(0.11*((rug/n)+(192.3*output*visc/port))**0.25)*dens*(port**2)/(output**5)
	return risultato

press_e = deqn(b,b)
press_1 = deqn(b,h1)
press_2 = deqn(b,h2)
press_3 = deqn(h1,h2)
press_4 = deqn(b,h3)
press_5 = deqn(b,h4)
press_6 = deqn(h3,h4)

sez_e = "{}x{}".format(str(b)[0:len(str(b))-2:1],str(b)[0:len(str(b))-2:1])
sez_1 = "{}x{}".format(str(b)[0:len(str(b))-2:1],str(h1)[0:len(str(h1))-2:1])
sez_2 = "{}x{}".format(str(b)[0:len(str(b))-2:1],str(h2)[0:len(str(h2))-2:1])
sez_3 = "{}x{}".format(str(h1)[0:len(str(h1))-2:1],str(h2)[0:len(str(h2))-2:1])
sez_4 = "{}x{}".format(str(b)[0:len(str(b))-2:1],str(h3)[0:len(str(h3))-2:1])
sez_5 = "{}x{}".format(str(b)[0:len(str(b))-2:1],str(h4)[0:len(str(h4))-2:1])
sez_6 = "{}x{}".format(str(h3)[0:len(str(h3))-2:1],str(h4)[0:len(str(h4))-2:1])

def vel(portata,base,altezza):
	outputv = (portata/3600.0)/((base/1000)*(altezza/1000))
	return outputv

vel_e = vel(port,b,b)
vel_1 = vel(port,b,h1)
vel_2 = vel(port,b,h2)
vel_3 = vel(port,h1,h2)
vel_4 = vel(port,b,h3)
vel_5 = vel(port,b,h4)
vel_6 = vel (port,h3,h4)

lista = [press_e,press_1,press_2,press_3,press_4,press_5,press_6]
sez_l = [sez_e,sez_1,sez_2,sez_3,sez_4,sez_5,sez_6]
vel_l = ["%.2f" % vel_e,"%.2f" % vel_1,"%.2f" % vel_2,"%.2f" % vel_3,"%.2f" % vel_4,"%.2f" % vel_5,"%.2f" % vel_6]
pres_l = ["%.4f" % press_e,"%.4f" % press_1,"%.4f" % press_2,"%.4f" % press_3,"%.4f" % press_4,"%.4f" % press_5,"%.4f" % press_6]
vel_lout =[vel_e,vel_1,vel_2,vel_3,vel_4,vel_5,vel_6]

Sez_g = []
Sez_s = []
Vel_g = []
Vel_s = []
Pres_g = []
Pres_s = []
Vel_out =[]

Sez_gi = None
Vel_gi = None
Pres_gi = None
Vel_outi =[]

for p,s,v,pp,vo in zip(lista,sez_l,vel_l,pres_l,vel_lout):
	if p > val_mi and p < val_ma:
		Sez_g.append(s)
		Vel_g.append(v)
		Pres_g.append(pp)
		Vel_out.append(vo)
	else:
		Sez_s.append(s)
		Vel_s.append(v)
		Pres_s.append(pp)
		Vel_outi.append(vo)


if len(Sez_g) == 0:
	Sez_g.append("Not Computed")
	Vel_out.append("n")

if Vel_out!= "n":
	for si,vi,pi,voi in zip(Sez_g,Vel_g,Pres_g,Vel_out):
		if voi == min(Vel_out):
			Sez_gi= si
			Vel_gi= vi
			Pres_gi= pi

from pyrevit import script

output = script.get_output()

if Vel_out!= "n":
	try:
		if min(Vel_out) <6:
			output.log_success('{} m/s NORMAL VALUE FOR CIVIL AND SILENT APPLICATIONS'.format(Vel_gi))
		elif min(Vel_out) >= 6  and min(Vel_out) < 14:
			output.log_info('{} m/s NORMAL VALUE FOR INDUSTRIAL APPLICATIONS'.format(Vel_gi))
		elif min(Vel_out)>=14 and min(Vel_out)<25:
			output.log_info('{} m/s NORMAL VALUE FOR SPECIAL APPLICATIONS'.format(Vel_gi))
		else:
			output.log_warning('{} m/s WARNING HIGH NOISE'.format(Vel_gi))
	except:
		pass

output.print_md(	'# **INPUT DATA:**')

print('\tRequired Flow Rate: {} m3/h\n\tDuct Material: {}\n\tAbsolute Roughness: {} mm\n\tGas Type: {} \n\tGas Density: {} kg/m3\n\tPressure Drop Min: {} mmH2O/m\n\tPressure Drop Max: {} mmH2O/m'.format(port,duct_m,rug,gas_t,dens,val_mi,val_ma))

output.print_md(	'# **CALCULATION RESULTS:**')

output.print_md(	'## **- EXACT VALUE (PRESSURE DROP VALUE IN RANGE AND MINIMUM SPEED):**')

print('---------------------------------------------------------------------------------\n\n\n')

print('\tExact Size:  {} mm\n\tPressure Drop:  {} mmH2O/m\n\tVelocity:  {} m/s\n\n\n'.format(Sez_gi,Pres_gi,Vel_gi))

output.print_md(	'## **- ALL OPTIMAL SIZES (PRESSURE DROP VALUE IN RANGE):**')

print('---------------------------------------------------------------------------------\n\n\n')

for i,p,v in zip (Sez_g,Pres_g,Vel_g):
	print('\tOptimal Sizes:  {} mm\n\tPressure Drop O.S:  {} mmH2O/m\n\tVelocity O.S:  {} m/s\n\n\n'.format(i,p,v))

output.print_md(	'## - ALTERNATIVE SIZES (PRESSURE DROP VALUE NEAR OF RANGE):')

print('---------------------------------------------------------------------------------\n\n\n')

for i,p,v in zip (Sez_s,Pres_s,Vel_s):
	print('\tAlternative Sizes:  {} mm\n\tPressure Drop A.S:  {} mmH2O/m\n\tVelocity A.S:  {} m/s\n\n\n'.format(i,p,v))

from docxtpl import DocxTemplate
from docx import Document
import pandas as pd 
import numpy as np 
from __pycache__ import numeros as num
import os

tipo = input('¿Que tipo de Resolución quiere hacer?\n1) Resolución General\n2) Resolución de Cuota\n\n(Presione 1 o 2): ')

#Defino el DataFrame con el que voy a trabajar#
df = pd.read_excel(r'C:\Users\Usuario\PyApps\Resolucion Auto\Templates\Resoluciones.xlsx')

#Con ésto elijo la fila del archivo Excel que quiero utilizar#
numero_de_fila = int(input('Número de fila: '))- 2

def resolucion(n):
	#Defino mis variables#
	global nombre, cuit, numero_actuacion, domicilio, monto_otorgado, monto_solicitado, monto_cuota_individual, cantidad_cuotas, considerando, monto, monto_art1, monto_art2, autorizados
	nombre = df['Institución'][n]
	cuit = df['C.U.I.T'][n]
	numero_actuacion = df['Nº de Actuación'][n]
	domicilio = df['Domicilio'][n]
	monto_otorgado = str(df['Monto Otorgado'][n])
	monto_solicitado = str(df['Monto Solicitado'][n])
	cantidad_cuotas = str(int(df['Cantidad de cuotas'][n]))
	lista = [df['Autorizado al Cobro Nº1'][n],df['Autorizado al Cobro Nº2'][n],df['Autorizado al Cobro Nº3'][n]]
	considerando = input("Considerando : ")

	#Defino mis funciones#
	def autorizados(n):
		def pronombre(x):
			if x == 'h':
				pronombre = 'al Sr. '
			else:
				pronombre = 'a la Sra. '
			return pronombre
		if lista[1].lower() == 'n':
			return '{} '.format(pronombre(df['Sexo 1'][n])) + str(lista[0])
		elif lista[-1].lower() == 'n':
			return '{} '.format(pronombre(df['Sexo 1'][n])) + str(lista[0]) + ', y/o {}'.format(pronombre(df['Sexo 2'][n])) + str(lista[1])
		else:
			return '{} '.format(pronombre(df['Sexo 1'][n])) + str(lista[0]) + ', {} '.format(pronombre(df['Sexo 2'][n])) + str(lista[1]) + ', y/o {} '.format(pronombre(df['Sexo 3'][n])) + str(lista[-1])
    
	autorizados = autorizados(n)

	def monto(n):
		if df['¿Cuotas?'][n] == "s":
			monto_cuota_individual = str(int(df['Monto Otorgado'][n]/df['Cantidad de cuotas'][n]))
			if monto_solicitado == monto_otorgado:
				return f'''{num.numero_a_moneda(monto_otorgado)} (${monto_otorgado[:-3]}.{monto_otorgado[-3:]},00), 
							a lo que esta instancia autoriza al pago en {cantidad_cuotas} cuotas de {num.numero_a_moneda(monto_cuota_individual)} 
							(${monto_cuota_individual[:-3]}.{monto_cuota_individual[-3:]},00) cada una, a partir de la fecha del presente instrumento legal
							'''
			else:
				return f'''{num.numero_a_moneda(monto_solicitado) }(${monto_solicitado[:-3]}.{monto_solicitado[-3:]},00), 
							a lo que esta instancia autoriza al pago por la suma total de {num.numero_a_moneda(monto_otorgado)} (${monto_otorgado[:-3]}.{monto_otorgado[-3:]},00), 
							pagadero en {cantidad_cuotas} cuotas de {num.numero_a_moneda(monto_cuota_individual)} (${monto_cuota_individual[:-3]}.{monto_cuota_individual[-3:]},00) cada una, 
							a partir de la fecha del presente instrumento legal
							'''
		else:
			if monto_solicitado == monto_otorgado:
				return f'{num.numero_a_moneda(monto_otorgado)} (${monto_otorgado[:-3]}.{monto_otorgado[-3:]},00)'
			else:
				return f'''{num.numero_a_moneda(monto_solicitado) }(${monto_solicitado[:-3]}.{monto_solicitado[-3:]},00), 
				a lo que esta instancia autoriza al pago por la suma total de {num.numero_a_moneda(monto_otorgado)} (${monto_otorgado[:-3]}.{monto_otorgado[-3:]},00)
				'''
    
	monto = monto(n)
    
	def monto_art1(n):
		if df['¿Cuotas?'][n] == "s":
			monto_cuota_individual = str(int(df['Monto Otorgado'][n]/df['Cantidad de cuotas'][n]))
			
			return f'''{num.numero_a_moneda(monto_otorgado).upper()} (${monto_otorgado[:-3]}.{monto_otorgado[-3:]},00), 
                        pagadero en {cantidad_cuotas} cuotas de {num.numero_a_moneda(monto_cuota_individual)} 
                        (${monto_cuota_individual[:-3]}.{monto_cuota_individual[-3:]},00) cada una
                        '''

		else:
			return f'{num.numero_a_moneda(monto_otorgado).upper()} (${monto_otorgado[:-3]}.{monto_otorgado[-3:]},00)'
	
	monto_art1 = monto_art1(n)
    
	def monto_art2(n):
		if df['¿Cuotas?'][n] == "s":
			monto_cuota_individual = str(int(df['Monto Otorgado'][n]/df['Cantidad de cuotas'][n]))
			return str(num.numero_a_moneda(monto_cuota_individual) + " ($" + str(format(int(monto_cuota_individual),'.2f')) + ") en concepto de primera cuota")
		else:                
			return str(num.numero_a_moneda(monto_otorgado).upper() + " ($" + str(format(int(monto_otorgado), '.2f')) + ")")		

	monto_art2 = monto_art2(n)


#Ejecuto mi función#
resolucion(numero_de_fila)

if tipo == '1':
	#resolucion general
	# #Abro mi template y agrego las variables al contexto#
	document = DocxTemplate(r'C:\Users\Usuario\PyApps\Resolucion Auto\Templates\TEMPLATE_GRAL.docx')
	
	context = {
	'nombre' : str(nombre),
	'cuit' : str(cuit),
	'numero_actuacion' : str(numero_actuacion),
	'domicilio' : str(domicilio),
	'monto_art1' : monto_art1,
	'monto_art2' : monto_art2,
	'monto' : str(monto),
	'cantidad_cuotas' : str(cantidad_cuotas),
	'autorizados' : str(autorizados),
	'considerando' : str(considerando)
	}

	document.render(context)

	if str(nombre)[0].upper() == 'F':
		path = (r'\\PRIVADASECRETAR\f\SecGral\APOYOS 2015-2016\APOYOS INSTITUCIONALES\FUNDACIONES\ '+ str(nombre).upper() +str(numero_actuacion[8:14])+'.docx')
	elif str(nombre)[0:2].upper() == 'CO':
		path = (r'\\PRIVADASECRETAR\f\SecGral\APOYOS 2015-2016\APOYOS INSTITUCIONALES\COOPERATIVAS\ '+ str(nombre).upper() +str(numero_actuacion[8:14])+'.docx')
	elif str(nombre)[0:2].upper() == 'CL':
		path = (r'\\PRIVADASECRETAR\f\SecGral\APOYOS 2015-2016\APOYOS INSTITUCIONALES\CLUB\ '+ str(nombre).upper() +str(numero_actuacion[8:14])+'.docx')
	else:
		path = (r'\\PRIVADASECRETAR\f\SecGral\APOYOS 2015-2016\APOYOS INSTITUCIONALES\ASOCIACIONES\ '+ str(nombre).upper() +str(numero_actuacion[8:14])+'.docx')

	document.save(path)
	os.startfile(path)


else:
	#cuota
	document = DocxTemplate(r'C:\Users\Usuario\PyApps\Resolucion Auto\Templates\TEMPLATE_CUOTA.docx')

	n_cuota = input("Número de cuota generada: ")
	if input('Decreto? (S/N): ').upper() == 'S':
		decreto = 'y Decreto ' + input('Número de Decreto: ')
	else:
		decreto = ''


	context = {
	'nombre' : str(nombre),
	'cuit' : str(cuit),
	'n_actuacion' : str(numero_actuacion),
	'decreto' : str(decreto),
	'domicilio' : str(domicilio),
	'monto_cuota' : str(num.numero_a_moneda(int(monto_otorgado)/int(cantidad_cuotas)).upper()) + ' (' + str(int(monto_otorgado)/int(cantidad_cuotas)) + ')',
	'n_cuota' : str(n_cuota),
	'n_cuota_rendida' : str((int(n_cuota) - 1)),
	'autorizados' : str(autorizados),
	'considerando' : str(considerando),
	'resol_gral' : input('Resolución General: '),
	'n_actuacion_rendida' : input('Número de la Actuación Rendida: ')
	}
	
	document.render(context)
	
	if str(nombre)[0].upper() == 'F':
		path = (r'\\PRIVADASECRETAR\f\SecGral\APOYOS 2015-2016\APOYOS INSTITUCIONALES\FUNDACIONES\ '+ str(nombre).upper() + ' (Cuota Nº'+str(n_cuota) + ') ' + str(numero_actuacion[8:14])+'.docx')
	elif str(nombre)[0:2].upper() == 'CO':
		path = (r'\\PRIVADASECRETAR\f\SecGral\APOYOS 2015-2016\APOYOS INSTITUCIONALES\COOPERATIVAS\ '+ str(nombre).upper() + ' (Cuota Nº'+str(n_cuota) + ') ' + str(numero_actuacion[8:14])+'.docx')
	elif str(nombre)[0:2].upper() == 'CL':
		path = (r'\\PRIVADASECRETAR\f\SecGral\APOYOS 2015-2016\APOYOS INSTITUCIONALES\CLUB\ '+ str(nombre).upper() + ' (Cuota Nº'+str(n_cuota) + ') ' + str(numero_actuacion[8:14])+'.docx')
	else:
		path = (r'\\PRIVADASECRETAR\f\SecGral\APOYOS 2015-2016\APOYOS INSTITUCIONALES\ASOCIACIONES\ '+ str(nombre).upper() + ' (Cuota Nº'+str(n_cuota) + ') ' + str(numero_actuacion[8:14])+'.docx')


	document.save(path)
	os.startfile(path)




from docxtpl import DocxTemplate
from docx import Document
import pandas as pd 
import numpy as np 
from __pycache__ import numeros as num
import os

#Defino el DataFrame con el que voy a trabajar#
df = pd.read_excel(r'C:\Users\mg_fa\OneDrive\Escritorio\Resolucion Auto\Templates\Resoluciones.xlsx')
#Con ésto elijo la fila del archivo Excel que quiero utilizar#
numero_de_fila = int(input('Número de fila: '))- 2
n = numero_de_fila

def resolucion(n):
	#Defino mis variables#
	global nombre, cuit, numero_actuacion, domicilio, monto_otorgado, cantidad_cuotas, considerando, autorizados
	nombre = df['Institución'][n]
	cuit = df['C.U.I.T'][n]
	numero_actuacion = df['Nº de Actuación'][n]
	domicilio = df['Domicilio'][n]
	monto_otorgado = df['Monto Otorgado'][n]
	cantidad_cuotas = df['Cantidad de cuotas'][n]
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

resolucion(n)

#cuota
resolucion = DocxTemplate(r'C:\Users\mg_fa\OneDrive\Escritorio\Resolucion Auto\Templates\TEMPLATE_CUOTA.docx')

if input('Decreto? (S/N): ').upper() == 'S':
    decreto = 'y Decreto ' + input('Número de Decreto: ')
else:
    decreto = ''


# ---------------------------- 1. GENERO LA NOTA ----------------------------------#
nota = DocxTemplate(r'C:\Users\Usuario\PyApps\Notas Auto\Templates\template nota prox cuota.docx')

numero_cuota_generada = input('Cuota/s Nº: ')
resolucion_gral = input('Número de Resolución: ')
numero_actuacion_rendida = input('Número de Actuación Rendida: ')

context_nota = { 
    'numero_actuacion' : df['Nº de Actuación'][n],
    'nombre_institucion' : df['Institución'][n],
    'numero_resol' : resolucion_gral, 
    'numero_cuota_generada' : 'Cuota Nº ' + str(numero_cuota_generada),
    }
context_nota['numero_actuacion_rendida'] = numero_actuacion_rendida
context_nota['numero_cuota_rendida'] = str(int(numero_cuota_generada) - 1)
context_nota['cuota_anterior'] = 'Cabe mencionar que la Actuación Simple Nº '+ context_nota['numero_actuacion_rendida'] +', por la cual la mencionada Institución efectuó la rendición de la Cuota Nº '+ context_nota['numero_cuota_rendida'] +' del aporte otorgado obra en el archivo de esta Secretaria sin observaciones por parte de la Unidad de Auditoría Interna.'

filename_nota = df['Institución'][n] + ' - Cuota N° ' + numero_cuota_generada + '.docx'

# path_nota = r'\\PRIVADASECRETAR\f\SecGral\APOYOS 2015-2016\MARTIN\PROVIDENCIAS\ ' + filename_nota
path_nota = r'C:\Users\Usuario\Desktop\test_nota.docx'
nota.render(context_nota)
nota.save(path_nota)

#--------------------------- 2. AHORA LA RESOLUCIÓN -----------------------------# 
monto_cuota = str(int(monto_otorgado/cantidad_cuotas))
context_resolucion = {
'nombre' : df['Institución'][n].upper(),
'cuit' : df['C.U.I.T'][n],
'numero_actuacion' : df['Nº de Actuación'][n],
'decreto' : decreto,
'domicilio' : df['Domicilio'][n],
'monto_cuota' : f'{num.numero_a_moneda(monto_cuota).upper()} (${str(monto_cuota)[:-3]}.{str(monto_cuota)[-3:]},00)',
'numero_cuota_generada' : numero_cuota_generada,
'numero_cuota_rendida' : str(int(numero_cuota_generada) - 1),
'autorizados' : str(autorizados),
'considerando' : str(considerando),
'resolucion_gral' : resolucion_gral,
'numero_actuacion_rendida' : numero_actuacion_rendida
}



# if str(nombre)[0].upper() == 'F':
#     path_resolucion = (r'\\PRIVADASECRETAR\f\SecGral\APOYOS 2015-2016\APOYOS INSTITUCIONALES\FUNDACIONES\ '+ nombre.upper() + ' (Cuota Nº'+ numero_cuota_generada + ') ' + numero_actuacion[8:14] +'.docx')
# elif str(nombre)[0:2].upper() == 'CO':
#     path_resolucion = (r'\\PRIVADASECRETAR\f\SecGral\APOYOS 2015-2016\APOYOS INSTITUCIONALES\COOPERATIVAS\ '+ nombre.upper() + ' (Cuota Nº'+ numero_cuota_generada + ') ' + numero_actuacion[8:14] +'.docx')
# elif str(nombre)[0:2].upper() == 'CL':
#     path_resolucion = (r'\\PRIVADASECRETAR\f\SecGral\APOYOS 2015-2016\APOYOS INSTITUCIONALES\CLUB\ '+ nombre.upper() + ' (Cuota Nº'+ numero_cuota_generada + ') ' + numero_actuacion[8:14] +'.docx')
# else:
#     path_resolucion = (r'\\PRIVADASECRETAR\f\SecGral\APOYOS 2015-2016\APOYOS INSTITUCIONALES\ASOCIACIONES\ '+ nombre.upper() + ' (Cuota Nº'+ numero_cuota_generada + ') ' + numero_actuacion[8:14] +'.docx')

path_resolucion = r'C:\Users\Usuario\Desktop\test.docx'

resolucion.render(context_resolucion)
resolucion.save(path_resolucion)

#-------------------------------------AL FINAL------------------------------------#
os.startfile(path_resolucion)
os.startfile(path_nota)
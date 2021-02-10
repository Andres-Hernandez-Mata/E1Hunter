"""
Uso: PEP, pip y APIs
Creado: Andrés Hernández Mata
Version: 1.0
Python: 3.8.2
Fecha: 10 de Febrero del 2021
"""


from pyhunter import PyHunter
from openpyxl import Workbook
import getpass


def Busqueda(organizacion):
    #Cantidad de resultados esperados de la búsqueda
    #El límite MENSUAL de Hunter es 50, cuidado!
    resultado = hunter.domain_search (company = organizacion, limit = 1, emails_type = 'personal')
    return resultado


def GuardarInformacion(datosEncontrados,organizacion):
    libro = Workbook()
    hoja = libro.create_sheet(organizacion)
    libro.save("Hunter" + organizacion + ".xlsx")
    #Agrega el codigo necesario para guardar en formato tabla
    #dentro del libro de Excel, información que consideres relevante
    #de lo obtenido en la búsqueda.
    libro.save("Hunter" + organizacion + ".xlsx")
    

print("Script para buscar información")
apikey = getpass.getpass("Ingresa tu API key: ")
hunter = PyHunter (apikey)
orga = input("Dominio a investigar: ")
datosEncontrados = Busqueda(orga)
if datosEncontrados == None:
    exit()
else:
    print(datosEncontrados)
    print(type(datosEncontrados))
    GuardarInformacion(datosEncontrados,orga)
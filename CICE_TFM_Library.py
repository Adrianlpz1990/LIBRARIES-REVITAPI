#-*- coding: utf-8 -*-

###COMMON LANGUAGE RUNTIME###	Es el entorno de ejecución que permite ejecutar código de varios lenguajes de programación distintos. Nos va a permitir ejecutar elementos que pertenecen a la API de Revit.
import clr		

###DOCUMENTMANAGER Y TRANSACTIONMANAGER###
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager			#Administracion del documento (Para localizar el documento)
from RevitServices.Transactions import TransactionManager		#Puntos de guardado (Para modificar el documento)
from System.Collections.Generic import *

###Importar RevitAPI###
clr.AddReference("RevitAPI")
import Autodesk
from Autodesk.Revit.DB import *
import Autodesk.Revit.DB as DB
from Autodesk.Revit.DB.Architecture import *
from Autodesk.Revit.DB.Structure import *
from Autodesk.Revit.DB.Plumbing import *
from Autodesk.Revit.DB.Electrical import *
from Autodesk.Revit.DB.Mechanical import *

###Importar RevitAPI USER INTERFACE###
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *

###GEOMETRIAS DE DYNAMO###
clr.AddReference('ProtoGeometry')
from Autodesk.DesignScript.Geometry import*

###NODOS DYNAMO###
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)			#Elementos espejo de Revit
clr.ImportExtensions(Revit.GeometryConversion)	#Unidades

###BIBLIOTECAS ESTANDAR DE IRONPYTHON###
import sys
sys.path.append(r'C:\Program Files (x86)\IronPython 2.7\Lib') 

###.NET###		Espacios de nombres del sistema de la raíz .NET. Permite traer listas fuertemente tipadas, matrices, colecciones genéricas e Icollection, 
import System		
from System import Array	#Listas
from System.Collections.Generic import *	#Listas fuertemente tipadas, etc

###SISTEMA OPERATIVO###
import os 

###EXCEPCIONES(ERRORES)###		Guardar excepciones
import traceback

###DOCUMENTO ACTUAL E INTERFACE DE USUARIO###
doc =  DocumentManager.Instance.CurrentDBDocument
app =  DocumentManager.Instance.CurrentUIApplication.Application
UIDocument =  DocumentManager.Instance.CurrentUIDocument
uiapp = DocumentManager.Instance.CurrentUIApplication
actuiapp = uiapp.ActiveUIDocument

########################################################################################################################################################

			#CONVERTIR UNIDADES#
#VERSION HASTA 2021#
def converToInt(param, valor):
	"""Permite convertir las unidades del parametro de las de proyecto a internas"""
	
	UIunit = param.DisplayUnitsType
	return UnitUtils.ConvertToInternalUnits(valor, UIunit)

def projectUnits(x):
	"""Lee el valor en unidades de proyecto"""

	UIunit = Document.GetUnits(doc).GetFormatOptions(UnitType.UT_Length).DisplayUnits
	return UnitUtils.ConvertFromInternalUnits(x, UIunit)

#VERSIÓN 2021 EN ADELANTE#
def converToInt2021(param, valor):
	"""Permite convertir las unidades del parametro de las de proyecto a internas"""

	UIunit = param.GetUnitTypeId()
	return UnitUtils.ConvertToInternalUnits(valor, UIunit)

#......................................................................................................

def tolist(x):
	"""Garantiza que el elemento sea iterable"""
	if hasattr(x,'__iter__'): return x
	else : return [x]

#......................................................................................................

def valorParametro(param):
	"""Da el valor de los parámetros sin importar el tipo que sean"""
	if param.StorageType == StorageType.String:
		return param.AsString()
	elif param.StorageType == StorageType.ElementId:
		return doc.GetElement(param.AsElementId())
	elif param.StorageType == StorageType.Double:
		return param.AsDouble()
	else:
		return param.AsInteger()

#......................................................................................................

def allValueParameters(x, param):
	"""Da el valor de los parámetros sin importar si el parametro esta completo o no"""
	try:
		value = valorParametro(x.LookupParameter(param))
		return value
	except:
		value = ""
		return value
	
#......................................................................................................

def currentSelection():
	"""Obtengo elementos seleccionados tanto del modelo, como del Project Browser"""
	def salida(x):
		if len(x) == 1: return x[0]
		else: return x
	selid = actuiapp.Selection.GetElementIds() #Ids seleccionados
	return salida([doc.GetElement(id).ToDSType(True) for id in selid])

#......................................................................................................

def GetRoomBoundaries(item):
	"""Obtengo las curvas y los puntos que definen el perimetro de una room"""
	doc = item.Document
	calculator = SpatialElementGeometryCalculator(doc)
	options = Autodesk.Revit.DB.SpatialElementBoundaryOptions()
	boundloc = Autodesk.Revit.DB.AreaVolumeSettings.GetAreaVolumeSettings(doc).GetSpatialElementBoundaryLocation(SpatialElementType.Room)
	options.SpatialElementBoundaryLocation = boundloc
	curvas = []
	try:
		for blist in item.GetBoundarySegments(options):
			points = []
			for b in blist:
				curve = b.GetCurve().ToProtoType()
				curvas.append(curve)
				if curve.StartPoint not in points:
					points.append(curve.StartPoint)
				elif curve.EndPoint not in points:
					points.append(curve.EndPoint)
				else:
					pass
	except:
		pass
	return curvas, points

#......................................................................................................


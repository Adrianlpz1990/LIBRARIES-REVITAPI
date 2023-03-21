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
import shutil		#Tratamiento de archivos,copiar, borrar, mover, etc
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

def createDictionary(keys, values):
	"""Crea un diccionario partiendo de dos listas"""
	return {keys[i] : values [i] for i in range(len(keys))}

#......................................................................................................

def setParameters(e, dic, param, key):
	"""Escribe el valor del parametro en el parametro del elemento obteniendo dicho valor de la clave de un diccionario"""
	try:
		e.LookupParameter(param).Set(dic[key])
	except:
		pass

#......................................................................................................

def unusedViewsCleanup(inicio, prefijo=':'):
    """Limpieza de vistas sin uso en planos. Toda vista que este fuera de los planos es candidata para ser eliminada. Tiene en cuenta vistas con dependencias de otras; vistas de llamada o Callout que dependen de otra vista; se eliminan las leyendas que no esten alojadas en planos. El usuario puede evitar la eliminacion de aquellas vistas que inicien con un determinado prefijo.
	Entrada:
	inicio ‹bool›: True para eliminar.
	prefijo ‹str›: Toda vista que inicie con ese prefijo se conserva. (Valor por defecto para prefijo dos puntos (caracter prohibido para nombres de vistas en Revit))
	Salida:
	Mensaje de exito o fallo ‹str>
	Pag318-Las mil y una funciones"""
    if inicio:
        idsConservar = set()
        planos = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).ToElements()
        for plano in planos:
            lista = plano.GetAllPlacedViews()
            for id in lista:
                idsConservar.add(id)
                vista = doc.GetElement(id)
                dependencia = vista.GetPrimaryViewId()
                if dependencia.IntegerValue != -1:
                    idsConservar.add(dependencia)
                parametro = BuiltInParameter.SECTION_PARENT_VIEW_NAME
                if vista.get_Parameter(parametro):
                    pariente = vista.get_Parameter(parametro).AsElementId()
                    if pariente.IntegerValue != -1:
                        idsConservar.add(pariente)
        
        idsVistas = set([vista.Id for vista in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).ToElements() if vista.IsTemplate is False and (vista.Name).startswith(prefijo) is False])
        idsVistasSinUso = idsVistas.difference(idsConservar)
        
        if doc.ActiveView.Id in idsVistasSinUso:
            salida = "La vista activa esta sin uso, vista activa para continuar."
        else:
            contador = 0
            if bool(idsVistasSinUso):
                TransactionManager.Instance.EnsureInTransaction(doc)
                for id in idsVistasSinUso:
                    try:
                        doc.Delete(id)
                        contador += 1
                    except:
                        pass
                TransactionManager.Instance.TransactionTaskDone()
                if contador == 0 and len(idsVistasSinUso) != 0:
                    salida = "No se ha eliminado ninguna vista y existen {} vistas sin uso.".format(len(idsVistasSinUso))
                elif contador != 0 and contador == len(idsVistasSinUso):
                    salida = "Se han eliminado el 100% de las vistas sin uso (Total: {} vistas).".format(contador)
                else:
                    salida = "Se han eliminado {} vistas de {} vistas sin uso.".format(contador, len(idsVistasSinUso))
            else:
                salida = "No hay vistas sin uso que eliminar."
    else:
        salida = "Necesita un True para iniciar la ejecución."
    return salida

#......................................................................................................

def seleccionPorNombreRegla(palabra=None):
	"""seleccionar una regla del asesor de rendimiento. Es necesario que se introduzca una palabra clave para localizar la regla.
	Entradas:
	Si el usuario no introduce ningun argumento extrae la lista completa de nombres para que pueda elegir.
	nombre <str>: Nombre de la regla.
	Salida:
	Si se introduce correctamente el nombre, se obtiene el id de la regla. Si se encuentran multiples opciones se pide mas letras.
	Pag336-Las mil y una funciones"""
	regla = []
	asesor = PerformanceAdviser.GetPerformanceAdviser()
	ids = asesor.GetAllRuleIds()
	nombres = [asesor.GetRuleName(id) for id in ids]
	
	if bool(palabra):
		for id, nombre in zip(ids, nombres):
			if palabra.lower() in nombre:
				regla.append(id)
			if regla:
				salida = (regla[0] if len(regla) == 1 else "Afinar busqueda, multiples opciones.")
			else:
				salida = "Ninguna coincidencia"
	else:
		salida = sorted(nombres)
	return salida

#......................................................................................................

def asesorEjecutarRegla(inicio, idRegla):
	"""Esta funcion ejecuta una regla del asesor de rendimiento y aporta un lista con todos los ids que deben ser revisados de la regla intoducida.
	Entrada:
	inicio ‹bool›: True para iniciar
	idRegla ‹PerformanceAdviserRuleId›: Regla a procesar
	Salida: Mensaje <String>
	Pag337-Las mil y una funciones"""
	if inicio:
		regla = List[PerformanceAdviserRuleId]([idRegla])
		ejecutar = PerformanceAdviser.GetPerformanceAdviser().ExecuteRules(doc, regla)
		if ejecutar:
			salida = ejecutar[0].GetFailingElements()
			adicionales = ejecutar[0].GetAdditionalElements()
			salida.AddRange(adicionales)
		else:
			salida = None
	else:
		salida = "Introducir un True para iniciar."
	return salida

#......................................................................................................

def unusedElementsCleanup(inicio):
	"""Esta funcion limpia todos los elementos sin uso del modelo, para realizar esto se usa el PerformanceAdviser que da acceso al listado completo de familias y tipos sin uso.
	Entrada:
	inicio <bool›: True para iniciar
	Salida: Mensaje <String>
	Dependencia:
	Funciones seleccionPorNombreRegla(), asesorEjecutarRegla()
	Pag338-Las mil y una funciones"""
	if inicio:
		regla = seleccionPorNombreRegla("unused")
		idsElementos = asesorEjecutarRegla(inicio, regla)
	
		if bool(idsElementos):
			try:
				TransactionManager.Instance.EnsureInTransaction(doc)
				doc.Delete(List[ElementId](idsElementos))
				TransactionManager.Instance.TransactionTaskDone()
				salida = "Elementos accesibles eliminados."
			except:
				salida = "Fallo en la ejecución."
		else:
			salida = ("Nada que eliminar.""\nEl asesor de rendimiento aporta \nuna lista vacia.")
	else:
		salida = "Introducir un True para iniciar."
	return salida

#......................................................................................................
def unusedSchedulesCleanup(inicio, prefijo=None):
	"""Uso:Limpieza de tablas sin uso en planos. Toda tabla que este fuera de los planos es candidata para ser eliminada. El usuario puede evitar la eliminacion de aquellas que inicien con un determinado prefijo.
	Entrada:
	inicio ‹bool›: True para eliminar.
	prefijo <str›: Toda tabla que inicie con ese prefijo se conserva.
	Valor nulo por defecto para prefijo
	Salida:
	Mensaje de exito o fallo <str>
	Pag320-Las mil y una funciones"""
	if inicio:
		idsConservar = set([tabla.ScheduleId for tabla in FilteredElementCollector(doc).OfClass(ScheduleSheetInstance).ToElements()])
		if prefijo:
			ids = set([tabla.Id for tabla in FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements() if not tabla.Name.startswith(prefijo)])
		else:
			ids = set([tabla.Id for tabla in FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()])
	
		idsTablasSinUso = ids.difference(idsConservar)
	
		if doc.ActiveView.Id.IntegerValue in idsTablasSinUso:
			salida = "La vista activa está en uso, es necesario cambiar de vista activa para continuar."
		else:
			contador = 0
			if idsTablasSinUso:
				TransactionManager.Instance.EnsureInTransaction(doc)
				for id in idsTablasSinUso:
					try:
						doc.Delete(id)
						contador += 1
					except:
						pass
				TransactionManager.Instance.TransactionTaskDone()
			
				if contador == 0 and len(idsTablasSinUso) != 0:
					salida = (("No se ha eliminado ninguna tabla \ny existen {} tablas sin uso.").format(len(idsTablasSinUso)))
				elif contador != 0 and contador == len(idsTablasSinUso):
					salida = (("Se han eliminado el 100% de las tablas \nsin uso (Total: {} tablas)").format(contador))
				else:
					salida = (("Se han eliminado {} tablas \nde {} tablas sin uso.").format(contador, len(idsTablasSinUso)))
			else:
				salida = "No hay tablas sin uso que eliminar."
	else:
		salida = "Necesita un True para iniciar \nla ejecución."
	return salida
#......................................................................................................

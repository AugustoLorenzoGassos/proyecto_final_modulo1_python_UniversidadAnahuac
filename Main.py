# -*- coding: utf-8 -*-
"""
Created on Thu Mar  5 13:54:38 2026

@author: Augusto Lorenzo Gassós

Proyecto Final: Módulo 1

Título: Sistema de Calificaciones y Asistencia (CLI) + Reporte 
Objetivo: crear una app de consola (menú) para registrar alumnos, calificaciones y 
asistencia; aplicar reglas con operadores lógicos, condicionales, ciclos, funciones, 
importaciones de la librería estándar y generar un reporte básico con pandas.

"""

import json
import os
import re
import datetime
import numpy as np
import time
import matplotlib.pyplot as plt
import pandas as pd
import csv

import Datos

Usuarios, Alumnos, Calificaciones, Asistencias = Datos.RecuperarDatos()

def Encabezado():
    print("\n")
    print("Sistema de Calificaciones y Asistencia\n".center(100))
    

def ValidarUsuario(NUsuario, CUsuario):
    #Validar cedenciales del usuario para inicar sesión
    UsuarioRegistrado = False
    for i in range(0,len(Usuarios["Usuarios"]),1):
        if Usuarios['Usuarios'][i]['Nombre'] == NUsuario and Usuarios['Usuarios'][i]['Contraseña'] == CUsuario:
            UsuarioRegistrado = True
            break
    return(UsuarioRegistrado, Usuarios['Usuarios'][i]['NombreCompleto'], Usuarios['Usuarios'][i]['Status'])

def IniciarSesion():
        
    OpcionSistema = 0
    while True:

        os.system("cls")
        Encabezado()
        print(f"Usuario: {NombreUsuarioCompleto}\n")
        print("Página principal".center(100))
        
        print("1: Alumnos\n2. Estadísticas\n3. Reporte de pandas\n4. Cerrar sesión\n")
        try:
            OpcionSistema = int(input("Selecciona la opción: "))
        except ValueError:
            OpcionSistema = 5
            
        match OpcionSistema:
            case 1:
                os.system(("cls"))
                MostrarAlumnos()
            case 2:
                os.system(("cls"))
                VerEstadisticas()
            case 3:
                os.system(("cls"))
                reporte_pandas()
            case 4:
                #with open(".\ListaAlumnos.json", "w") as f:
                #    json.dump(Alumnos, f)            
                #with open(".\ListaCalificaciones.json", "w") as f:
                #    json.dump(Calificaciones, f)            
                Datos.GuadarDatos(Alumnos, Calificaciones, Asistencias)
                os.system("cls")
                break
            case _:
                print("\nOpción no reconocida\n")
                Continuar = input("¿Desea intentar nuevamente? (S/N) ")
                if Continuar.lower() == "s":
                    os.system("cls")
                    continue
                else:
                    break
                    
                

def MostrarAlumnos():
    
    while True:
        
        os.system("cls")
        Encabezado()
        print(f"Usuario: {NombreUsuarioCompleto}\n")
        print("Lista de alumnos\n".center(100))
        
        # Mostrar la lista
        print(f"{'Matrícula':^20}{'Nombre completo':^50}{'Grupo':^10}{'Correo electrónico':^30}{'Alta':^20}{'Status':^20}")
        for i in range(0,len(Alumnos["Alumnos"]),1):
            DescripcionStatus = ""
            match Alumnos['Alumnos'][i]['Status']:
                case 1:
                    DescripcionStatus = "Aprobado"
                case 2:
                    DescripcionStatus = "En riesgo"
                case 3:
                    DescripcionStatus = "Reprpobado"
                case _:
                    DescripcionStatus = "Sin calcular"
            print(f"{Alumnos['Alumnos'][i]['Matricula']:<20}{Alumnos['Alumnos'][i]['NombreCompleto']:<50}{Alumnos['Alumnos'][i]['Grupo']:^10}{Alumnos['Alumnos'][i]['CorreoElectronico']:<30}{Alumnos['Alumnos'][i]['FechaIngreso']:^20}{DescripcionStatus:<20}")
        print("\n")
        
        OpcionAlumnos = 0
        print("1. Alta de alumnos\n2. Buscar alumno\n3. Exportar datos a Excel\n4. Salir de alumnos\n")
        
        try:
            OpcionAlumnos = int(input("Selecciona la opción: "))
        except ValueError:
            OpcionAlumnos = 5
        
        match OpcionAlumnos:
            case 1:
                AgregarAlumno()
            case 2:
                os.system("cls")
                Encabezado()
                print(f"Usuario: {NombreUsuarioCompleto}\n")
                print("Buscar alumno\n".center(100))
                BuscarMatricula = input("Escribe la matrícula del alumno a buscar: ")    
                BuscarAlumno(BuscarMatricula)
            case 3:
                os.system("cls")
                Encabezado()
                ExportarCSV()
            case 4:
                os.system("cls")
                break
            case _:
                print("\nOpción no reconocida\n")
                Continuar = input("¿Desea intentar nuevamente? (S/N) ")
                if Continuar.lower() == "s":
                    os.system("cls")
                    continue
                else:
                    break

def AgregarAlumno():

    FechaLimiteRegistro = "30/11/2025"
    ContinuarRegistro = ""
    
    while True:
        
        os.system("cls")
        Encabezado()
        print(f"Usuario: {NombreUsuarioCompleto}\n")
        print("Alta de alumnos\n".center(100))
        
        LlenadoCorrecto = True
        
        NAlumno = input("Escribe el nombre del alumno: ")
        MAlumno = input("Escribe la matrícula del alumno: ")
        GAlumno  = input("Escribe el grupo del alumno: ")
        CEAlumno = input("Escribe el correo electrónico del alumno: ")
        FIAlumno = input("Escribe la fecha de ingreso (dd/mm/YYYY): ")
    
        if len(NAlumno) == 0 or len(MAlumno)==0 or len(GAlumno) == 0 or len(CEAlumno) == 0 or len(FIAlumno)==0:
            print("No puedes dejar campos vacíos. Revisa la información...")
            LlenadoCorrecto = False
        else:
            ExpresionCorreo = re.compile(r'([A-Za-z0-9]+[.-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+')    
            ExpresionMatricula = re.compile(r'[A-Z]{4,}[0-9]{7,}')
            if not re.match(ExpresionCorreo, CEAlumno):
                print("\nEl formarto para el correo electrónico es incorrecto")
                LlenadoCorrecto = False
            if not re.match(ExpresionMatricula, MAlumno):
                print("\nEl formato para la matrícula es incorrecto")
                LlenadoCorrecto = False
            for i in range(0,len(Alumnos["Alumnos"]),1):
                if Alumnos['Alumnos'][i]['Matricula'] == MAlumno:
                    print("\nLa matrícula ya existe en la base de datos")
                    LlenadoCorrecto = False
            try:
                FechaTransformada = datetime.datetime.strptime(FIAlumno, "%d/%m/%Y").date()
                if datetime.datetime.strptime(FIAlumno, "%d/%m/%Y") > datetime.datetime.strptime(FechaLimiteRegistro, "%d/%m/%Y"):
                    print("\nLa fecha de ingreso rebasa la fecha límite para el registro")
                    LlenadoCorrecto = False
            except ValueError:
                print("\nFormato de fecha incorrecto")
                LlenadoCorrecto = False
    
        if LlenadoCorrecto == True:
            print("\nInformación correcra.\n")
            ContinuarRegistro = input("¿Deseas guardar el registro? (S/N) ")
            if ContinuarRegistro.lower()=="s":
                #print(datetime.datetime.strptime(FIAlumno, "%d/%m/%Y").date())
                Alumnos["Alumnos"].append({"Matricula":MAlumno, "NombreCompleto": NAlumno, "Grupo": GAlumno, "CorreoElectronico": CEAlumno, "FechaIngreso": FIAlumno, "Status": 0})
                Calificaciones["Calificaciones"].append({"Matricula":MAlumno, "Espanol":0.0, "Matematicas":0.0, "Naturales":0.0, "Geografia":0.0, "Civismo":0.0, "PromedioFinal":0.0})
                Asistencias["Asistencias"].append({"Matricula":MAlumno,"Faltas":0,"Asistencias":0})
                #Asistencias
            os.system("cls")
            break
        else:
            ContinuarRegistro = input("\n¿Deseas volver a capturar la información? (S/N) ")
            if ContinuarRegistro.lower()=="s":
                os.system("cls")
                Encabezado()
                continue
            else:
                os.system("cls")
                break
            

def BuscarAlumno(PMatricula):

    while True:
        
        os.system("cls")
        Encabezado()
        print(f"Usuario: {NombreUsuarioCompleto}\n")
        #print("Buscar alumno\n".center(100))
        
        AlumnoEncontrado = False
        
        for i in range(0,len(Alumnos["Alumnos"]),1):
            if Alumnos['Alumnos'][i]['Matricula'] == PMatricula:
                AlumnoEncontrado = True
                break
    
        if AlumnoEncontrado:

            DescripcionStatus = ""
            match Alumnos['Alumnos'][i]['Status']:
                case 1:
                    DescripcionStatus = "Aprobado"
                case 2:
                    DescripcionStatus = "En riesgo"
                case 3:
                    DescripcionStatus = "Reprpobado"
                case _:
                    DescripcionStatus = "Sin calcular"

            print("\nDatos del alumno".center((100)))
            print("\n")
            print(f"{'Matrícula':^20}{'Nombre completo':^50}{'Grupo':^10}{'Correo electrónico':^30}{'Alta':^20}{'Status':^20}")
            print(f"{Alumnos['Alumnos'][i]['Matricula']:<20}{Alumnos['Alumnos'][i]['NombreCompleto']:<50}{Alumnos['Alumnos'][i]['Grupo']:^10}{Alumnos['Alumnos'][i]['CorreoElectronico']:<30}{Alumnos['Alumnos'][i]['FechaIngreso']:^20}{DescripcionStatus:<20}")
            print("\n")
            print("1. Registrar calificaciones\n2. Registro de asistencia\n3. Validar estatus\n4. Ver gráficas\n5. Regresar a la lista de alumnos")
            
            try:
                OpcionAlumnoIndividual = int(input("\nSelecciona la opción: "))
            except ValueError:
                OpcionAlumnoIndividual = 6
                
            match OpcionAlumnoIndividual:
                case 1:
                    os.system("cls")
                    Encabezado()
                    RegistrarCalificaciones(PMatricula)
                case 2:
                    os.system("cls")
                    Encabezado()
                    RegistrarAsistencias(PMatricula)
                case 3:
                    os.system("cls")
                    Encabezado()
                    CalcularEstatus(PMatricula)
                case 4:
                    os.system("cls")
                    Encabezado()
                    VerGraficas(PMatricula)
                case 5:
                    os.system("cls")
                    break
                case _:
                    Continuar = input("\n¿Deseas intentar nuevamente (S/N)? ")
                    if Continuar.lower()=="s":
                        continue
                    else:
                        os.system("cls")
                        break
        else:
            print("\n")
            print(f"La matrícula {PMatricula} no esta registrada en la base de datos\n")
            ContinuarBusqueda = input("¿Deseas intentar nuevamente? (S/N)? ")
            if ContinuarBusqueda.lower() == "s":
                continue
            else:
                break
   
def RegistrarCalificaciones(AMatricula):
    
    print("Registro de calificaciones\n".center(100))
    os.system("cls")
    
    CalificacionesEncontradas = False
    ItemCalificaciones = 0
    ModificarCalificacione =  ""
    CalificacionEspanol = 0
    CalificacionMatematicas = 0
    CalificacionNaturales = 0
    CalificacionGeografia = 0
    CalificacionCivismo = 0
    for i in range(0,len(Calificaciones["Calificaciones"]),1):
        if Calificaciones['Calificaciones'][i]['Matricula'] == AMatricula:
            CalificacionesEncontradas = True
            ItemCalificaciones = i
            break
    
    if CalificacionesEncontradas == True:
        print("Español: ",Calificaciones['Calificaciones'][ItemCalificaciones]['Espanol'])
        print("Matemáticas: ",Calificaciones['Calificaciones'][ItemCalificaciones]['Matematicas'])
        print("Ciencias Naturales: ",Calificaciones['Calificaciones'][ItemCalificaciones]['Naturales'])
        print("Geografía: ",Calificaciones['Calificaciones'][ItemCalificaciones]['Geografia'])
        print("Educación Cívica: ",Calificaciones['Calificaciones'][ItemCalificaciones]['Civismo'])
        print("\n")
        print("Promedio final: ", Calificaciones['Calificaciones'][ItemCalificaciones]['PromedioFinal'])
        print("\nEl alumno ya tiene registradas calificaciones")
        ModificarCalificacione = input("\n¿Deseas modificar las calificaciones registradas S/N?  (Al concluir de capturar las calificaciones finales se calculará el promedio final) ")
        if ModificarCalificacione.lower() == "s":
            ErrorCalificaciones = False
            print("\n")
            CalificacionEspanol = float(input("Calificación de español: "))
            CalificacionMatematicas = float(input("Calificación de matemáticas: "))
            CalificacionNaturales = float(input("Calificación de ciencias naturales: "))
            CalificacionGeografia = float(input("Calificación de geografía: "))
            CalificacionCivismo = float(input("Calificación de educación cívica: "))
            #Validar calificaciones
            print("\n")
            if (CalificacionEspanol<0 or CalificacionEspanol>100):
                print("La calificación de español esta fuera del rango")
                ErrorCalificaciones = True
            if (CalificacionMatematicas<0 or CalificacionMatematicas>100):
                print("La calificación de matemáticas esta fuera del rango")
                ErrorCalificaciones = True
            if (CalificacionNaturales<0 or CalificacionNaturales>100):
                print("La calificación de ciencias naturales esta fuera del rango")
                ErrorCalificaciones = True
            if (CalificacionGeografia<0 or CalificacionGeografia>100):
                print("La calificación de geografía esta fuera del rango")
                ErrorCalificaciones = True
            if (CalificacionCivismo<0 or CalificacionCivismo>100):
                print("La calificación de educación cívica esta fuera del rango")
                ErrorCalificaciones = True
            
            if ErrorCalificaciones == False:
                Calificaciones['Calificaciones'][ItemCalificaciones]['Espanol']=CalificacionEspanol  
                Calificaciones['Calificaciones'][ItemCalificaciones]['Matematicas']=CalificacionMatematicas
                Calificaciones['Calificaciones'][ItemCalificaciones]['Naturales']=CalificacionNaturales
                Calificaciones['Calificaciones'][ItemCalificaciones]['Geografia']=CalificacionGeografia
                Calificaciones['Calificaciones'][ItemCalificaciones]['Civismo']=CalificacionCivismo
                CalificacionesPromedio = []
                CalificacionesPromedio.append(CalificacionEspanol)
                CalificacionesPromedio.append(CalificacionMatematicas)
                CalificacionesPromedio.append(CalificacionNaturales)
                CalificacionesPromedio.append(CalificacionGeografia)
                CalificacionesPromedio.append(CalificacionCivismo)
                Calificaciones['Calificaciones'][ItemCalificaciones]['PromedioFinal'] = np.mean(CalificacionesPromedio)
            else:
                print("\nLas calificaciones no fueron almacenadas ya que se presentaron errores en su captura")
                input("\nPresione enter para continuar ...")
    
def RegistrarAsistencias(AMatricula):
    print("Registro de asistencias\n".center(100))
    os.system("cls")

    AsistenciasEncontradas = False
    ItemAsistencias = 0
    CantidadAsistencias = 0
    CantidadFaltas = 0
    ErrorRegistroAsistencias = False
    
    for i in range(0,len(Asistencias["Asistencias"]),1):
        if Asistencias['Asistencias'][i]['Matricula'] == AMatricula:
            AsistenciasEncontradas = True     
            ItemAsistencias = i
            break
    if AsistenciasEncontradas == True:
        print("Asistencias: ",Asistencias['Asistencias'][ItemAsistencias]['Asistencias'])
        print("Faltas: ",Asistencias['Asistencias'][ItemAsistencias]['Faltas'])
        print("\nEl alumno ya tiene registradas sus asistencias")
        ModificarAsistencias = input("\n¿Deseas modificar las asistencias registradas S/N?: ")
        print("\n")
        if ModificarAsistencias.lower()=="s":
            try:
                CantidadAsistencias = int(input("Cantidad de asistencias: "))
                CantidadFaltas = int(input("Cantidad de faltas: "))
            except ValueError:
                    print("/nLas asistencias y faltas deben de ser un número\nEspere 10 segundos para continuar")
                    CantidadAsistencias = 0
                    CantidadFaltas = 0
                    time.sleep((10))
                    ErrorRegistroAsistencias = True
            
            if ErrorRegistroAsistencias == False:
                Asistencias['Asistencias'][ItemAsistencias]['Asistencias']=CantidadAsistencias
                Asistencias['Asistencias'][ItemAsistencias]['Faltas']=CantidadFaltas
            
def CalcularEstatus(AMatricula):
    
    CalificacionesEncontradas = False
    AsistenciasEncontradas = False
    
    for i in range(0,len(Calificaciones["Calificaciones"]),1):
        if Calificaciones['Calificaciones'][i]['Matricula'] == AMatricula:
            PromedioFinal = Calificaciones['Calificaciones'][i]['PromedioFinal']
            CalificacionesEncontradas = True
            print("Promedio final: ", PromedioFinal)
            break
    for i in range(0,len(Asistencias["Asistencias"]),1):
        if Asistencias['Asistencias'][i]['Matricula'] == AMatricula:
            CantidadAsistencias = int(Asistencias['Asistencias'][i]['Asistencias'])
            CantidadFaltas = int(Asistencias['Asistencias'][i]['Faltas'])
            AsistenciasEncontradas = True     
            print("Asistencias: ", CantidadAsistencias)
            print("Faltas: ", CantidadFaltas)
            print("Porcentaje de asistencia: ", (CantidadAsistencias/(CantidadAsistencias+CantidadFaltas))*100)
            break
        
    if CalificacionesEncontradas == True and AsistenciasEncontradas == True:
        if PromedioFinal >= 70 and (CantidadAsistencias/(CantidadAsistencias+CantidadFaltas))*100>=80:
            StatusAlumno = 1
        if PromedioFinal < 70 or (CantidadAsistencias/(CantidadAsistencias+CantidadFaltas))*100<80:
            StatusAlumno = 2
        if PromedioFinal < 50 or (CantidadAsistencias/(CantidadAsistencias+CantidadFaltas))*100<60:
            StatusAlumno = 3
        
        match StatusAlumno:
            case 1:
                print("\nStatus del alumno: Aprobado")
            case 2:
                print("\nStatus del alumno: En riesgo")
            case 3:
                print("\nStatus del alumno: Reprobado")
            
        for i in range(0,len(Alumnos["Alumnos"]),1):
            if Alumnos['Alumnos'][i]['Matricula'] == AMatricula:
                Alumnos['Alumnos'][i]['Status'] = StatusAlumno 
                break
            
        input("\nPresiona enter para continuar...")
    else:
        if CalificacionesEncontradas == False:
            print("No e encunetra el registro de calificaciones\n")
        if AsistenciasEncontradas == False:
            print("No se encuentra el registro de asistencia\n")
        print("No se pudo calcular el status del alumno.\nEsepra 10 segundos para continuar.")
        time.sleep(10)
        
def VerGraficas(AMatricula):
    print("Gráficas\n".center(100))
    
    CalificacionesGrafica = []
    Materias = ["Esp.", "Mat.", "Nat.","Geo.","Civ","Prom."]
    
    FaltasAsistencias = []
    
    #Obtener calificaciones
    for i in range(0,len(Calificaciones["Calificaciones"]),1):
        if Calificaciones['Calificaciones'][i]['Matricula'] == AMatricula:
            CalificacionesGrafica.append(Calificaciones['Calificaciones'][i]['Espanol'])
            CalificacionesGrafica.append(Calificaciones['Calificaciones'][i]['Matematicas'])
            CalificacionesGrafica.append(Calificaciones['Calificaciones'][i]['Naturales'])
            CalificacionesGrafica.append(Calificaciones['Calificaciones'][i]['Geografia'])
            CalificacionesGrafica.append(Calificaciones['Calificaciones'][i]['Civismo'])
            CalificacionesGrafica.append(Calificaciones['Calificaciones'][i]['PromedioFinal'])
            break
    
    #Obtener matemáticas
    for i in range(0,len(Asistencias["Asistencias"]),1):
        if Asistencias['Asistencias'][i]['Matricula'] == AMatricula:
            FaltasAsistencias.append(Asistencias['Asistencias'][i]['Asistencias'])
            FaltasAsistencias.append(Asistencias['Asistencias'][i]['Faltas'])
    
    plt.Figure(layout="constrained")
    
    if len(CalificacionesGrafica) > 0:
        plt.subplot(221)
        plt.bar(Materias, CalificacionesGrafica, color='lightgreen')
        plt.xlabel("Materias")
        plt.ylabel("Calificaciones")
        plt.title("Registro de calificaciones")        
    
    if len(FaltasAsistencias) > 0:
        plt.subplot(222)
        plt.pie(FaltasAsistencias, labels=["Asist.","Faltas"])
        plt.title("Registro de asistencias")
        
    plt.grid(True)
    plt.show()
    input("\nPresiona enter para continuar ...")

def ExportarCSV():
    
    DiccionarioCalificacionesLinea = {}
    
    ListaCalificaciones = Calificaciones["Calificaciones"]
    Calificacionesframe = pd.DataFrame(ListaCalificaciones)
        
    #Método 1
    Calificacionesframe.to_csv("Calificaciones1.csv", index=False)
    Calificacionesframe.to_excel("Calificaciones1.xlsx", index=False)
    
    #Método 2
    Contador = 0
    with open('Calificacines2.csv', mode='w', newline='') as f:
        Campos = ["Matricula","Espanol","Matematicas","Naturales","Geografia","Civismo","PromedioFinal"]
        writer = csv.DictWriter(f, fieldnames=Campos)
        writer.writeheader()
        
        for i in range(0,len(Calificaciones["Calificaciones"]),1):
            DiccionarioCalificacionesLinea = Calificaciones["Calificaciones"][i]
            writer.writerow(DiccionarioCalificacionesLinea)    
            Contador += 1
    
    print("El proceso de exportae a Excel (CSV) finalizó correctamente\n")
    print(f"Se exportó un total de  {Contador} registgros.\n")
    input("Presiona enter para continuar...")
  
def VerEstadisticas():
    
    os.system("cls")
    Encabezado()
    print(f"Usuario: {NombreUsuarioCompleto}\n")
    print("Estadísticas\n".center(100))
    
    ListaCalificaciones = Calificaciones["Calificaciones"]
    Calificacionesframe = pd.DataFrame(ListaCalificaciones)
    ListaAlumnos = Alumnos["Alumnos"]
    AlumnosFrame = pd.DataFrame(ListaAlumnos)
    
    print("Datos originales:\n\n", Calificacionesframe)
    print("\n")
    print("Promedios por materia")
    print(f"Español: {Calificacionesframe['Espanol'].mean():.2f} ")
    print(f"Matemátias: {Calificacionesframe['Matematicas'].mean():.2f}")
    print(f"Ciencias Naturales: {Calificacionesframe['Naturales'].mean():.2f}")
    print(f"Geografía: {Calificacionesframe['Geografia'].mean():.2f}")
    print(f"Educación Cívica: {Calificacionesframe['Civismo'].mean():.2f}")
    print(f"Promedio Final: {Calificacionesframe['PromedioFinal'].mean():.2f}")

    print("\n")
    print("Mediana por materia\n")
    print(f"Español: {Calificacionesframe['Espanol'].median():.2f} ")
    print(f"Matemátias: {Calificacionesframe['Matematicas'].median():.2f}")
    print(f"Ciencias Naturales: {Calificacionesframe['Naturales'].median():.2f}")
    print(f"Geografía: {Calificacionesframe['Geografia'].median():.2f}")
    print(f"Educación Cívica: {Calificacionesframe['Civismo'].median():.2f}")
    print(f"Promedio Final: {Calificacionesframe['PromedioFinal'].median():.2f}")
    
    print("\n")
    print("Alumnos por grupo\n")
    print(AlumnosFrame.groupby(["Grupo"]).count())
    
    #Valores ordenados
    print("\nCalificaciones ordenadas por promedio final")
    CalificacionesframeOrdenado = Calificacionesframe.sort_values(by="PromedioFinal",ascending=False)
    print(CalificacionesframeOrdenado)

    #Outliers en español
    print("\nIdentificando oitliers en la materia de español\n")
    Media = Calificacionesframe["Espanol"].mean()
    DesviacionEstadar = Calificacionesframe["Espanol"].std()
    Umbral = 1
    
    Calificacionesframe['Es_Outlier'] = np.abs(Calificacionesframe['Espanol'] - Media) > (Umbral * DesviacionEstadar)
    print(Calificacionesframe)
    
    print("\nEstadísticas descriptivas\n")
    Estadisticas = Calificacionesframe.describe()
    print(Estadisticas)
    ExportarCSV = input("¿Deseas exportar las estadísticas descriptivas a jun archivo CSV (S/N)? ")
    if ExportarCSV.lower() == "s":
        Calificacionesframe.describe().to_csv("EstadisticasDescriptivas.csv")
    
def reporte_pandas():
    os.system("cls")
    Encabezado()

    ListaAlumnos = Alumnos["Alumnos"]
    AlumnosFrame = pd.DataFrame(ListaAlumnos)
    ListaCalificaciones = Calificaciones["Calificaciones"]
    Calificacionesframe = pd.DataFrame(ListaCalificaciones)

    print(f"Usuario: {NombreUsuarioCompleto}\n")
    print("Repoorte con pandas\n".center(100))
    print("Alumnos head")
    print(AlumnosFrame.head())
    print("\nAlumnos describe")
    print(AlumnosFrame.describe())
    print("\nAlumnos por value_counts Grupo")
    print(AlumnosFrame.value_counts("Grupo"))
    print("\nPrimeros 5 en promedio final")
    print(Calificacionesframe["PromedioFinal"].nlargest(n=5))
    input("\nPresiona enter para continuar...")

#Inicio del sistema
NombreUsuario = ""
Contraseña = ""
Opcion = 0
Continuar = ""

while True:
     
    ContinuarInicio = ""
    
    os.system("cls")
    Encabezado()
    
    print("Inicio de sesión".center((100)))
    print("1. Inicio de sesión\n2. Salir\n")
    
    try:
        Opcion = int(input("Introduzca la opción seleccionada: "))
    except ValueError:
        Opcion = 3
    
    match Opcion:
        case 1:
            print("\n")
            NombreUsuario = input("Usuario: ")
            Contraseña = input("Contraseña: ")
            if len(NombreUsuario)>0 and len(Contraseña)>0:
                Encontrado, NombreUsuarioCompleto, StatusUsuario = ValidarUsuario(NombreUsuario, Contraseña)
                if Encontrado == True and StatusUsuario == "A":
                    os.system("cls")
                    IniciarSesion()
                elif Encontrado == True and StatusUsuario == "I":
                    print("Advertencia: Usuario inactivo")
                else:
                    print("\nAdvertencia: Usuario no registrado.\n")
                    ContinuarInicio = input("¿Desea intentar nuevamente (S/N)? ")
                    if ContinuarInicio.lower() == "s":
                        continue
                    else:
                        break
            else:
                print("\nEl usuario y/o contraseña no pueden estar vacíos\n")
                ContinuarInicio = input("¿Desea intentar nuevamente (S/N)? ")
                if ContinuarInicio.lower() == "s":
                    continue
                else:
                    break
        case 2:
            os.system("cls")
            break
        case _:
            print("\nOpción no reconocida\n")
            Continuar = input("¿Desea intentar nuevamente? (S/N) ")
            if Continuar.lower() == "s":
                os.system("cls")
                continue
            else:
                break

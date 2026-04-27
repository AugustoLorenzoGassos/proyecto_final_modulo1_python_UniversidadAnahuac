# -*- coding: utf-8 -*-
"""
Created on Sat Mar  7 20:07:54 2026

@author: Augusto Lorenzo Gassós

"""

import os
import json

Ruta = os.path.join(os.getcwd(), "Datos")
ArchivoUsuarios = os.path.join(Ruta, "ListaUsuarios.Json")
ArchivoAlumnos = os.path.join(Ruta, "ListaAlumnos.Json")
ArchivoCalificaciones = os.path.join(Ruta, "ListaCalificaciones.Json")
ArchivoAsistencias = os.path.join(Ruta, "ListaAsistencias.Json")

Usuarios = {}
Alumnos = {}
Calificaciones = {}
Asistencias = {}

def RecuperarDatos():
    #Usuarios
    with open(ArchivoUsuarios, "r") as f:
        Usuarios = json.load(f)
    #Alumnos
    with open(ArchivoAlumnos, "r") as f:
        Alumnos = json.load(f)
    #Califiaciones
    with open(ArchivoCalificaciones, "r") as f:
        Calificaciones = json.load(f)
    #Asistencias
    with open(ArchivoAsistencias, "r") as f:
        Asistencias = json.load(f)
    
    return Usuarios, Alumnos, Calificaciones, Asistencias

def GuadarDatos(DatosAlumnos, DatosCalificaciones, DatosAsistencias):
    with open(ArchivoAlumnos, "w") as f:
        json.dump(DatosAlumnos, f)         
    
    with open(ArchivoCalificaciones, "w") as f:
        json.dump(DatosCalificaciones, f)         

    with open(ArchivoAsistencias, "w") as f:
        json.dump(DatosAsistencias, f)         

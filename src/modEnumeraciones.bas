Attribute VB_Name = "modEnumeraciones"
Option Compare Database
Option Explicit



' Enumeración de roles de usuario para el sistema CONDOR
' Define los diferentes tipos de roles disponibles
Public Enum UserRole
    RolDesconocido = 0
    RolAdmin = 1
    RolCalidad = 2
    RolTecnico = 3
End Enum

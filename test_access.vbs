Option Explicit

Dim objAccess, dbPath, password

dbPath = "C:\Proyectos\cli\Expedientes.accdb"
password = "dpddpd"

WScript.Echo "Probando apertura de Access..."
WScript.Echo "Ruta: " & dbPath
WScript.Echo "Contraseña: " & password

On Error Resume Next

' Crear instancia de Access
Set objAccess = CreateObject("Access.Application")
If Err.Number <> 0 Then
    WScript.Echo "ERROR: No se pudo crear instancia de Access: " & Err.Description & " (Código: " & Err.Number & ")"
    WScript.Quit 1
End If

WScript.Echo "Instancia de Access creada exitosamente"

' Configurar Access
objAccess.Visible = False
objAccess.UserControl = False

' Abrir la base de datos
WScript.Echo "Intentando abrir base de datos..."
objAccess.OpenCurrentDatabase dbPath, False, password

If Err.Number <> 0 Then
    WScript.Echo "ERROR: No se pudo abrir la base de datos: " & Err.Description & " (Código: " & Err.Number & ")"
    objAccess.Quit
    Set objAccess = Nothing
    WScript.Quit 1
End If

WScript.Echo "Base de datos abierta exitosamente!"

' Verificar que se abrió correctamente
If objAccess.CurrentProject Is Nothing Then
    WScript.Echo "ERROR: CurrentProject es Nothing"
    objAccess.Quit
    Set objAccess = Nothing
    WScript.Quit 1
End If

WScript.Echo "CurrentProject disponible: " & objAccess.CurrentProject.Name

' Cerrar Access
objAccess.Quit
Set objAccess = Nothing

WScript.Echo "Prueba completada exitosamente!"
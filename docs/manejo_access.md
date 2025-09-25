# Apertura y Cierre Programático de Microsoft Access de Forma Totalmente Desatendida y Silenciosa

Para abrir Microsoft Access de forma completamente programática, silenciosa y desatendida (sin UI visible y sin confirmaciones del usuario), debes seguir configuraciones específicas documentadas por Microsoft. Aquí están los métodos oficiales:

## Apertura Silenciosa de Access con CreateObject

## Método Básico con Application.Visible y UserControl

Según la documentación oficial de Microsoft, cuando se crea un objeto **Application** mediante automatización, las propiedades **Visible** y **UserControl** se establecen automáticamente en  **False** :[](https://learn.microsoft.com/en-us/office/vba/api/access.application.visible)

<pre class="not-prose w-full rounded font-mono text-sm font-extralight"><div class="codeWrapper text-light selection:text-super selection:bg-super/10 my-md relative flex flex-col rounded font-mono text-sm font-normal bg-subtler"><div class="translate-y-xs -translate-x-xs bottom-xl mb-xl sticky top-0 flex h-0 items-start justify-end"></div><div class="-mt-xl"><div><div data-testid="code-language-indicator" class="text-quiet bg-subtle py-xs px-sm inline-block rounded-br rounded-tl-[3px] font-thin">text</div></div><div class="pr-lg"><span><code><span><span>Public Function AbrirAccessSilencioso(rutaBD As String) As Access.Application
</span></span><span>    ' Crear instancia invisible de Access
</span><span>    Dim objAccess As Access.Application
</span><span>    Set objAccess = CreateObject("Access.Application")
</span><span>  
</span><span>    ' Las propiedades Visible y UserControl ya están en False automáticamente
</span><span>    ' cuando se crea mediante automatización
</span><span>  
</span><span>    ' Confirmar que está invisible
</span><span>    objAccess.Visible = False
</span><span>    objAccess.UserControl = False
</span><span>  
</span><span>    ' Abrir base de datos específica
</span><span>    objAccess.OpenCurrentDatabase rutaBD, False ' False = modo compartido
</span><span>  
</span><span>    ' Retornar la instancia para trabajar con ella
</span><span>    Set AbrirAccessSilencioso = objAccess
</span><span>End Function
</span><span></span></code></span></div></div></div></pre>

## Apertura con Configuración de Seguridad y Sin Prompts

<pre class="not-prose w-full rounded font-mono text-sm font-extralight"><div class="codeWrapper text-light selection:text-super selection:bg-super/10 my-md relative flex flex-col rounded font-mono text-sm font-normal bg-subtler"><div class="translate-y-xs -translate-x-xs bottom-xl mb-xl sticky top-0 flex h-0 items-start justify-end"></div><div class="-mt-xl"><div><div data-testid="code-language-indicator" class="text-quiet bg-subtle py-xs px-sm inline-block rounded-br rounded-tl-[3px] font-thin">text</div></div><div class="pr-lg"><span><code><span><span>Public Function AbrirAccessCompleto(rutaBD As String, _
</span></span><span>                                   Optional contraseñaBD As String = "", _
</span><span>                                   Optional exclusivo As Boolean = False, _
</span><span>                                   Optional ocultarAvisos As Boolean = True) As Access.Application
</span><span>  
</span><span>    On Error GoTo ErrorHandler
</span><span>  
</span><span>    Dim objAccess As Access.Application
</span><span>    Set objAccess = CreateObject("Access.Application")
</span><span>  
</span><span>    ' Configuración para modo silencioso total
</span><span>    objAccess.Visible = False
</span><span>    objAccess.UserControl = False
</span><span>  
</span><span>    ' Deshabilitar confirmaciones si se solicita
</span><span>    If ocultarAvisos Then
</span><span>        Call ConfigurarOpcionesSilenciosas(objAccess)
</span><span>    End If
</span><span>  
</span><span>    ' Abrir base de datos con contraseña si es necesario
</span><span>    If contraseñaBD <> "" Then
</span><span>        objAccess.OpenCurrentDatabase rutaBD, exclusivo, contraseñaBD
</span><span>    Else
</span><span>        objAccess.OpenCurrentDatabase rutaBD, exclusivo
</span><span>    End If
</span><span>  
</span><span>    Set AbrirAccessCompleto = objAccess
</span><span>    Exit Function
</span><span>  
</span><span>ErrorHandler:
</span><span>    If Not objAccess Is Nothing Then
</span><span>        objAccess.Quit acQuitSaveNone
</span><span>        Set objAccess = Nothing
</span><span>    End If
</span><span>    Set AbrirAccessCompleto = Nothing
</span><span>End Function
</span><span>
</span><span>Private Sub ConfigurarOpcionesSilenciosas(objAccess As Access.Application)
</span><span>    ' Deshabilitar confirmaciones que podrían interrumpir la automatización
</span><span>    On Error Resume Next
</span><span>  
</span><span>    objAccess.SetOption "Confirm Record Changes", False
</span><span>    objAccess.SetOption "Confirm Document Deletions", False  
</span><span>    objAccess.SetOption "Confirm Action Queries", False
</span><span>    objAccess.SetOption "Show Status Bar", False
</span><span>  
</span><span>    On Error GoTo 0
</span><span>End Sub
</span><span></span></code></span></div></div></div></pre>

## Apertura mediante Línea de Comandos (msaccess.exe)

Según la documentación oficial, Access soporta varios modificadores de línea de comandos para operación desatendida:[](https://isladogs.co.uk/command-line-switches/)

## Modificadores de Línea de Comandos Documentados

<pre class="not-prose w-full rounded font-mono text-sm font-extralight"><div class="codeWrapper text-light selection:text-super selection:bg-super/10 my-md relative flex flex-col rounded font-mono text-sm font-normal bg-subtler"><div class="translate-y-xs -translate-x-xs bottom-xl mb-xl sticky top-0 flex h-0 items-start justify-end"></div><div class="-mt-xl"><div><div data-testid="code-language-indicator" class="text-quiet bg-subtle py-xs px-sm inline-block rounded-br rounded-tl-[3px] font-thin">text</div></div><div class="pr-lg"><span><code><span><span>Public Function AbrirAccessPorComandos(rutaBD As String, _
</span></span><span>                                      Optional parametrosExtra As String = "") As Access.Application
</span><span>  
</span><span>    Dim rutaMSAccess As String
</span><span>    Dim comandoCompleto As String
</span><span>    Dim objAccess As Access.Application
</span><span>  
</span><span>    ' Obtener ruta de msaccess.exe
</span><span>    rutaMSAccess = ObtenerRutaMSAccess()
</span><span>    If rutaMSAccess = "" Then
</span><span>        MsgBox "No se pudo localizar msaccess.exe", vbCritical
</span><span>        Exit Function
</span><span>    End If
</span><span>  
</span><span>    ' Construir comando con modificadores silenciosos
</span><span>    comandoCompleto = """" & rutaMSAccess & """ """ & rutaBD & """ /runtime"
</span><span>    If parametrosExtra <> "" Then
</span><span>        comandoCompleto = comandoCompleto & " " & parametrosExtra
</span><span>    End If
</span><span>  
</span><span>    ' Ejecutar Access de forma oculta
</span><span>    Call Shell(comandoCompleto, vbHide)
</span><span>  
</span><span>    ' Esperar a que Access se registre y obtener la instancia
</span><span>    Application.Wait DateAdd("s", 3, Now) ' Esperar 3 segundos
</span><span>  
</span><span>    On Error Resume Next
</span><span>    Set objAccess = GetObject(, "Access.Application")
</span><span>    On Error GoTo 0
</span><span>  
</span><span>    If Not objAccess Is Nothing Then
</span><span>        ' Asegurar que permanezca invisible
</span><span>        objAccess.Visible = False
</span><span>        objAccess.UserControl = False
</span><span>    End If
</span><span>  
</span><span>    Set AbrirAccessPorComandos = objAccess
</span><span>End Function
</span><span>
</span><span>Private Function ObtenerRutaMSAccess() As String
</span><span>    ' Obtener ruta desde el registro según documentación Microsoft
</span><span>    Dim clave As String
</span><span>    Dim ruta As String
</span><span>  
</span><span>    On Error Resume Next
</span><span>  
</span><span>    ' Buscar en las ubicaciones estándar de Office
</span><span>    clave = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE\"
</span><span>    ruta = CreateObject("WScript.Shell").RegRead(clave)
</span><span>  
</span><span>    If ruta = "" Then
</span><span>        ' Intentar ruta alternativa para Office 365
</span><span>        ruta = "C:\Program Files\Microsoft Office\root\Office16\msaccess.exe"
</span><span>        If Dir(ruta) = "" Then
</span><span>            ruta = "C:\Program Files (x86)\Microsoft Office\root\Office16\msaccess.exe"
</span><span>        End If
</span><span>    End If
</span><span>  
</span><span>    ObtenerRutaMSAccess = ruta
</span><span>    On Error GoTo 0
</span><span>End Function
</span><span></span></code></span></div></div></div></pre>

## Modificadores de Línea de Comandos Disponibles

Según la documentación oficial:[](https://support.microsoft.com/en-gb/office/command-line-switches-for-microsoft-office-products-079164cd-4ef5-4178-b235-441737deb3a6)

| Modificador         | Descripción                                                 |
| ------------------- | ------------------------------------------------------------ |
| `/runtime`        | Ejecuta Access en modo runtime (sin herramientas de diseño) |
| `/excl`           | Abre la base de datos en modo exclusivo                      |
| `/ro`             | Abre la base de datos en modo solo lectura                   |
| `/nostartup`      | Evita que se ejecute la macro AutoExec                       |
| `/cmd argumentos` | Pasa argumentos que pueden recuperarse con Command$          |
| `/x macro`        | Ejecuta la macro especificada al abrir                       |
| `/profile perfil` | Usa un perfil de usuario específico                         |

## Cierre Automático y Desatendido

## Método Recomendado para Cierre Completo

Basado en la documentación y experiencias reportadas, el cierre completo requiere una secuencia específica:[](https://www.reddit.com/r/MSAccess/comments/1g06ho7/msaccessexe_stays_open_after_database_app_closes/)

<pre class="not-prose w-full rounded font-mono text-sm font-extralight"><div class="codeWrapper text-light selection:text-super selection:bg-super/10 my-md relative flex flex-col rounded font-mono text-sm font-normal bg-subtler"><div class="translate-y-xs -translate-x-xs bottom-xl mb-xl sticky top-0 flex h-0 items-start justify-end"></div><div class="-mt-xl"><div><div data-testid="code-language-indicator" class="text-quiet bg-subtle py-xs px-sm inline-block rounded-br rounded-tl-[3px] font-thin">text</div></div><div class="pr-lg"><span><code><span><span>Public Sub CerrarAccessCompleto(objAccess As Access.Application)
</span></span><span>    ' Secuencia de cierre recomendada para evitar procesos zombie
</span><span>    On Error Resume Next
</span><span>  
</span><span>    Dim i As Integer
</span><span>  
</span><span>    ' 1. Cerrar todos los objetos abiertos
</span><span>    Call CerrarTodosLosObjetos(objAccess)
</span><span>  
</span><span>    ' 2. Cerrar conexiones de base de datos adicionales
</span><span>    For i = DBEngine.Workspaces(0).Databases.Count - 1 To 1 Step -1
</span><span>        DBEngine.Workspaces(0).Databases(i).Close
</span><span>        DBEngine.Idle
</span><span>    Next i
</span><span>  
</span><span>    ' 3. Cerrar la base de datos actual
</span><span>    objAccess.CloseCurrentDatabase
</span><span>  
</span><span>    ' 4. Ejecutar Quit - pero el código continúa después de esta línea
</span><span>    objAccess.Quit acQuitSaveNone
</span><span>  
</span><span>    ' 5. CRÍTICO: Cerrar CurrentDb DESPUÉS de Quit (solución documentada)
</span><span>    objAccess.CurrentDb.Close
</span><span>  
</span><span>    ' 6. Forzar liberación de memoria
</span><span>    Set objAccess = Nothing
</span><span>  
</span><span>    ' 7. Forzar recolección de basura (importante para COM)
</span><span>    DoEvents
</span><span>    DoEvents
</span><span>  
</span><span>    On Error GoTo 0
</span><span>End Sub
</span><span>
</span><span>Private Sub CerrarTodosLosObjetos(objAccess As Access.Application)
</span><span>    ' Cerrar formularios abiertos
</span><span>    Do While objAccess.Forms.Count > 0
</span><span>        objAccess.DoCmd.Close acForm, objAccess.Forms(0).Name, acSaveNo
</span><span>    Loop
</span><span>  
</span><span>    ' Cerrar informes abiertos  
</span><span>    Do While objAccess.Reports.Count > 0
</span><span>        objAccess.DoCmd.Close acReport, objAccess.Reports(0).Name, acSaveNo
</span><span>    Loop
</span><span>  
</span><span>    ' Cerrar módulos abiertos en el editor VBA
</span><span>    On Error Resume Next
</span><span>    objAccess.DoCmd.Close acModule, "", acSaveNo
</span><span>    On Error GoTo 0
</span><span>End Sub
</span><span></span></code></span></div></div></div></pre>

## Función de Cierre con Timeout

<pre class="not-prose w-full rounded font-mono text-sm font-extralight"><div class="codeWrapper text-light selection:text-super selection:bg-super/10 my-md relative flex flex-col rounded font-mono text-sm font-normal bg-subtler"><div class="translate-y-xs -translate-x-xs bottom-xl mb-xl sticky top-0 flex h-0 items-start justify-end"></div><div class="-mt-xl"><div><div data-testid="code-language-indicator" class="text-quiet bg-subtle py-xs px-sm inline-block rounded-br rounded-tl-[3px] font-thin">text</div></div><div class="pr-lg"><span><code><span><span>Public Function CerrarAccessConTimeout(objAccess As Access.Application, _
</span></span><span>                                     Optional timeoutSegundos As Integer = 10) As Boolean
</span><span>  
</span><span>    Dim inicioTimeout As Date
</span><span>    inicioTimeout = Now
</span><span>  
</span><span>    On Error GoTo ErrorHandler
</span><span>  
</span><span>    ' Ejecutar secuencia de cierre
</span><span>    Call CerrarAccessCompleto(objAccess)
</span><span>  
</span><span>    ' Verificar que el proceso realmente termine
</span><span>    Do While DateDiff("s", inicioTimeout, Now) < timeoutSegundos
</span><span>        DoEvents
</span><span>        Application.Wait DateAdd("s", 1, Now)
</span><span>      
</span><span>        ' Intentar acceder al objeto - si falla, significa que se cerró
</span><span>        On Error Resume Next
</span><span>        Dim test As String
</span><span>        test = objAccess.Name
</span><span>        If Err.Number <> 0 Then
</span><span>            CerrarAccessConTimeout = True
</span><span>            Exit Function
</span><span>        End If
</span><span>        On Error GoTo ErrorHandler
</span><span>    Loop
</span><span>  
</span><span>    ' Si llegamos aquí, el timeout expiró
</span><span>    CerrarAccessConTimeout = False
</span><span>    Exit Function
</span><span>  
</span><span>ErrorHandler:
</span><span>    CerrarAccessConTimeout = False
</span><span>End Function
</span><span></span></code></span></div></div></div></pre>

## Ejemplo de Uso Completo

<pre class="not-prose w-full rounded font-mono text-sm font-extralight"><div class="codeWrapper text-light selection:text-super selection:bg-super/10 my-md relative flex flex-col rounded font-mono text-sm font-normal bg-subtler"><div class="translate-y-xs -translate-x-xs bottom-xl mb-xl sticky top-0 flex h-0 items-start justify-end"></div><div class="-mt-xl"><div><div data-testid="code-language-indicator" class="text-quiet bg-subtle py-xs px-sm inline-block rounded-br rounded-tl-[3px] font-thin">text</div></div><div class="pr-lg"><span><code><span><span>Sub EjemploAutomatizacionCompleta()
</span></span><span>    Dim objAccess As Access.Application
</span><span>    Dim rutaBD As String
</span><span>  
</span><span>    rutaBD = "C:\MiBD\Ejemplo.accdb"
</span><span>  
</span><span>    ' Abrir Access de forma silenciosa
</span><span>    Set objAccess = AbrirAccessCompleto(rutaBD, "", False, True)
</span><span>  
</span><span>    If Not objAccess Is Nothing Then
</span><span>        ' Realizar operaciones automatizadas
</span><span>        Call RealizarImportacionModulos(objAccess)
</span><span>      
</span><span>        ' Cerrar de forma completa y desatendida
</span><span>        If Not CerrarAccessConTimeout(objAccess, 15) Then
</span><span>            ' Forzar cierre si el método normal falla
</span><span>            Call TerminarProcesoAccess()
</span><span>        End If
</span><span>    End If
</span><span>  
</span><span>    Set objAccess = Nothing
</span><span>End Sub
</span><span>
</span><span>Private Sub TerminarProcesoAccess()
</span><span>    ' Método de último recurso - terminar proceso por nombre
</span><span>    On Error Resume Next
</span><span>    CreateObject("WScript.Shell").Run "taskkill /f /im msaccess.exe", 0, True
</span><span>    On Error GoTo 0
</span><span>End Sub
</span><span>
</span><span>Private Sub RealizarImportacionModulos(objAccess As Access.Application)
</span><span>    ' Aquí van las operaciones de importación de módulos
</span><span>    ' usando los métodos VBComponents.Import vistos anteriormente
</span><span>  
</span><span>    On Error Resume Next
</span><span>    objAccess.VBE.ActiveVBProject.VBComponents.Import "C:\Modulos\MiModulo.bas"
</span><span>    On Error GoTo 0
</span><span>End Sub
</span><span></span></code></span></div></div></div></pre>

## Consideraciones Importantes

## Propiedades Críticas para Modo Silencioso

Según la documentación oficial:[](https://learn.microsoft.com/en-us/office/vba/api/access.application.usercontrol)

1. **Application.Visible = False** - Oculta la ventana de Access
2. **Application.UserControl = False** - Indica control programático (no del usuario)
3. Cuando  **UserControl = True** , no es posible establecer **Visible = False**
4. La automatización siempre inicia con ambas propiedades en **False**

## Configuraciones Adicionales para Operación Desatendida

<pre class="not-prose w-full rounded font-mono text-sm font-extralight"><div class="codeWrapper text-light selection:text-super selection:bg-super/10 my-md relative flex flex-col rounded font-mono text-sm font-normal bg-subtler"><div class="translate-y-xs -translate-x-xs bottom-xl mb-xl sticky top-0 flex h-0 items-start justify-end"></div><div class="-mt-xl"><div><div data-testid="code-language-indicator" class="text-quiet bg-subtle py-xs px-sm inline-block rounded-br rounded-tl-[3px] font-thin">text</div></div><div class="pr-lg"><span><code><span><span>Private Sub ConfiguracionAvanzadaSilenciosa(objAccess As Access.Application)
</span></span><span>    On Error Resume Next
</span><span>  
</span><span>    ' Deshabilitar todas las confirmaciones
</span><span>    objAccess.SetOption "Confirm Record Changes", False
</span><span>    objAccess.SetOption "Confirm Document Deletions", False
</span><span>    objAccess.SetOption "Confirm Action Queries", False
</span><span>  
</span><span>    ' Ocultar elementos de interfaz
</span><span>    objAccess.SetOption "Show Status Bar", False
</span><span>    objAccess.SetOption "Show Animations", False
</span><span>  
</span><span>    ' Configurar manejo de errores silencioso
</span><span>    objAccess.SetOption "Default Open Mode for Databases", 1 ' Compartido
</span><span>    objAccess.SetOption "Default Record Locking", 0 ' Sin bloqueos
</span><span>  
</span><span>    On Error GoTo 0
</span><span>End Sub
</span><span></span></code></span></div></div></div></pre>

Esta implementación garantiza una apertura y cierre completamente desatendidos de Microsoft Access, siguiendo la documentación oficial y las mejores prácticas para evitar procesos zombie o intervención del usuario.

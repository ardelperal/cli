



# Ejemplo Completo: Creación de Formulario con Controles de Forma Programática en Access

A continuación se presenta un ejemplo completo que demuestra cómo crear un formulario con varios controles desde cero de forma totalmente programática y desatendida, desde la apertura de la base de datos hasta su cierre perfecto.

## Código Principal - Automatización Completa

<pre class="not-prose w-full rounded font-mono text-sm font-extralight"><div class="codeWrapper text-light selection:text-super selection:bg-super/10 my-md relative flex flex-col rounded font-mono text-sm font-normal bg-subtler"><div class="translate-y-xs -translate-x-xs bottom-xl mb-xl sticky top-0 flex h-0 items-start justify-end"></div><div class="-mt-xl"><div><div data-testid="code-language-indicator" class="text-quiet bg-subtle py-xs px-sm inline-block rounded-br rounded-tl-[3px] font-thin">text</div></div><div class="pr-lg"><span><code><span><span>Public Sub EjemploCreacionFormularioCompleto()
</span></span><span>    ' Ejemplo completo de automatización: abrir Access, crear formulario con controles,
</span><span>    ' guardarlo y cerrar todo de forma desatendida
</span><span>  
</span><span>    Dim objAccess As Access.Application
</span><span>    Dim rutaBD As String
</span><span>  
</span><span>    rutaBD = "C:\MiBaseDatos\Ejemplo.accdb"
</span><span>  
</span><span>    ' 1. Abrir Access de forma silenciosa
</span><span>    Set objAccess = AbrirAccessSilencioso(rutaBD)
</span><span>  
</span><span>    If Not objAccess Is Nothing Then
</span><span>        ' 2. Crear el formulario con controles
</span><span>        Call CrearFormularioConControles(objAccess)
</span><span>      
</span><span>        ' 3. Cerrar Access de forma completa
</span><span>        Call CerrarAccessCompleto(objAccess)
</span><span>    End If
</span><span>  
</span><span>    Set objAccess = Nothing
</span><span>    MsgBox "Proceso completado exitosamente", vbInformation
</span><span>End Sub
</span><span>
</span><span>Private Function AbrirAccessSilencioso(rutaBD As String) As Access.Application
</span><span>    ' Abrir Access de forma invisible y desatendida
</span><span>    On Error GoTo ErrorHandler
</span><span>  
</span><span>    Dim objAccess As Access.Application
</span><span>    Set objAccess = CreateObject("Access.Application")
</span><span>  
</span><span>    ' Configurar modo silencioso
</span><span>    objAccess.Visible = False
</span><span>    objAccess.UserControl = False
</span><span>  
</span><span>    ' Deshabilitar confirmaciones
</span><span>    objAccess.SetOption "Confirm Record Changes", False
</span><span>    objAccess.SetOption "Confirm Document Deletions", False
</span><span>    objAccess.SetOption "Confirm Action Queries", False
</span><span>    objAccess.SetOption "Show Status Bar", False
</span><span>  
</span><span>    ' Abrir base de datos
</span><span>    objAccess.OpenCurrentDatabase rutaBD, False
</span><span>  
</span><span>    Set AbrirAccessSilencioso = objAccess
</span><span>    Exit Function
</span><span>  
</span><span>ErrorHandler:
</span><span>    If Not objAccess Is Nothing Then
</span><span>        objAccess.Quit acQuitSaveNone
</span><span>        Set objAccess = Nothing
</span><span>    End If
</span><span>    Set AbrirAccessSilencioso = Nothing
</span><span>End Function
</span><span></span></code></span></div></div></div></pre>

## Función de Creación de Formulario con Controles

<pre class="not-prose w-full rounded font-mono text-sm font-extralight"><div class="codeWrapper text-light selection:text-super selection:bg-super/10 my-md relative flex flex-col rounded font-mono text-sm font-normal bg-subtler"><div class="translate-y-xs -translate-x-xs bottom-xl mb-xl sticky top-0 flex h-0 items-start justify-end"></div><div class="-mt-xl"><div><div data-testid="code-language-indicator" class="text-quiet bg-subtle py-xs px-sm inline-block rounded-br rounded-tl-[3px] font-thin">text</div></div><div class="pr-lg"><span><code><span><span>Private Sub CrearFormularioConControles(objAccess As Access.Application)
</span></span><span>    ' Crear un formulario completo con múltiples controles programáticamente
</span><span>    On Error GoTo ErrorHandler
</span><span>  
</span><span>    Dim frm As Access.Form
</span><span>    Dim ctlLabel As Access.Control
</span><span>    Dim ctlTextBox As Access.Control
</span><span>    Dim ctlComboBox As Access.Control
</span><span>    Dim ctlCheckBox As Access.Control
</span><span>    Dim ctlCommandButton As Access.Control
</span><span>    Dim ctlDatePicker As Access.Control
</span><span>  
</span><span>    Dim nombreFormulario As String
</span><span>    Dim x As Long, y As Long, anchoControl As Long, altoControl As Long
</span><span>  
</span><span>    ' Crear el formulario base
</span><span>    Set frm = objAccess.CreateForm()
</span><span>    nombreFormulario = frm.Name
</span><span>  
</span><span>    Debug.Print "Formulario creado: " & nombreFormulario
</span><span>  
</span><span>    ' Configurar propiedades del formulario
</span><span>    With frm
</span><span>        .Caption = "Formulario Creado Programáticamente"
</span><span>        .RecordSource = "" ' Sin origen de datos por ahora
</span><span>        .ScrollBars = 0 ' Sin barras de desplazamiento
</span><span>        .NavigationButtons = False
</span><span>        .RecordSelectors = False
</span><span>        .DividingLines = False
</span><span>        .AutoCenter = True
</span><span>        .PopUp = False
</span><span>        .Modal = False
</span><span>        .Width = 8000 ' Ancho en twips
</span><span>        .Section(acDetail).Height = 6000 ' Alto de la sección detalle
</span><span>        .Section(acHeader).Height = 800 ' Alto de encabezado
</span><span>        .Section(acHeader).Visible = True
</span><span>    End With
</span><span>  
</span><span>    ' Configurar posiciones iniciales
</span><span>    x = 500 ' Posición X inicial
</span><span>    y = 200 ' Posición Y inicial
</span><span>    anchoControl = 2000 ' Ancho estándar de controles
</span><span>    altoControl = 300 ' Alto estándar de controles
</span><span>  
</span><span>    ' 1. CREAR ETIQUETA DE TÍTULO EN EL ENCABEZADO
</span><span>    Set ctlLabel = objAccess.CreateControl(nombreFormulario, acLabel, acHeader, "", "", _
</span><span>                                         2000, 100, 4000, 400)
</span><span>    With ctlLabel
</span><span>        .Caption = "FORMULARIO DE EJEMPLO"
</span><span>        .FontName = "Arial"
</span><span>        .FontSize = 14
</span><span>        .FontBold = True
</span><span>        .TextAlign = 2 ' Centrado
</span><span>        .ForeColor = RGB(0, 0, 128) ' Azul oscuro
</span><span>    End With
</span><span>  
</span><span>    ' 2. CREAR CAMPO DE TEXTO CON ETIQUETA (Nombre)
</span><span>    Set ctlLabel = objAccess.CreateControl(nombreFormulario, acLabel, acDetail, "", "", _
</span><span>                                         x, y, anchoControl, altoControl)
</span><span>    ctlLabel.Caption = "Nombre:"
</span><span>    ctlLabel.FontBold = True
</span><span>  
</span><span>    Set ctlTextBox = objAccess.CreateControl(nombreFormulario, acTextBox, acDetail, "", "", _
</span><span>                                           x + anchoControl + 200, y, anchoControl + 500, altoControl)
</span><span>    With ctlTextBox
</span><span>        .Name = "txtNombre"
</span><span>        .FontName = "Arial"
</span><span>        .FontSize = 10
</span><span>    End With
</span><span>  
</span><span>    y = y + altoControl + 300 ' Siguiente línea
</span><span>  
</span><span>    ' 3. CREAR CAMPO DE EMAIL CON ETIQUETA
</span><span>    Set ctlLabel = objAccess.CreateControl(nombreFormulario, acLabel, acDetail, "", "", _
</span><span>                                         x, y, anchoControl, altoControl)
</span><span>    ctlLabel.Caption = "Email:"
</span><span>    ctlLabel.FontBold = True
</span><span>  
</span><span>    Set ctlTextBox = objAccess.CreateControl(nombreFormulario, acTextBox, acDetail, "", "", _
</span><span>                                           x + anchoControl + 200, y, anchoControl + 500, altoControl)
</span><span>    With ctlTextBox
</span><span>        .Name = "txtEmail"
</span><span>        .FontName = "Arial"
</span><span>        .FontSize = 10
</span><span>        .InputMask = ""
</span><span>    End With
</span><span>  
</span><span>    y = y + altoControl + 300 ' Siguiente línea
</span><span>  
</span><span>    ' 4. CREAR COMBO BOX CON ETIQUETA (Categoría)
</span><span>    Set ctlLabel = objAccess.CreateControl(nombreFormulario, acLabel, acDetail, "", "", _
</span><span>                                         x, y, anchoControl, altoControl)
</span><span>    ctlLabel.Caption = "Categoría:"
</span><span>    ctlLabel.FontBold = True
</span><span>  
</span><span>    Set ctlComboBox = objAccess.CreateControl(nombreFormulario, acComboBox, acDetail, "", "", _
</span><span>                                            x + anchoControl + 200, y, anchoControl + 500, altoControl)
</span><span>    With ctlComboBox
</span><span>        .Name = "cboCategoria"
</span><span>        .FontName = "Arial"
</span><span>        .FontSize = 10
</span><span>        .RowSourceType = "Value List"
</span><span>        .RowSource = "Cliente;Proveedor;Empleado;Otro"
</span><span>        .LimitToList = True
</span><span>    End With
</span><span>  
</span><span>    y = y + altoControl + 300 ' Siguiente línea
</span><span>  
</span><span>    ' 5. CREAR CAMPO DE FECHA CON ETIQUETA
</span><span>    Set ctlLabel = objAccess.CreateControl(nombreFormulario, acLabel, acDetail, "", "", _
</span><span>                                         x, y, anchoControl, altoControl)
</span><span>    ctlLabel.Caption = "Fecha Registro:"
</span><span>    ctlLabel.FontBold = True
</span><span>  
</span><span>    Set ctlDatePicker = objAccess.CreateControl(nombreFormulario, acTextBox, acDetail, "", "", _
</span><span>                                              x + anchoControl + 200, y, anchoControl, altoControl)
</span><span>    With ctlDatePicker
</span><span>        .Name = "txtFechaRegistro"
</span><span>        .FontName = "Arial"
</span><span>        .FontSize = 10
</span><span>        .Format = "Short Date"
</span><span>        .ShowDatePicker = 1 ' Mostrar selector de fecha
</span><span>    End With
</span><span>  
</span><span>    y = y + altoControl + 300 ' Siguiente línea
</span><span>  
</span><span>    ' 6. CREAR CHECKBOX
</span><span>    Set ctlCheckBox = objAccess.CreateControl(nombreFormulario, acCheckBox, acDetail, "", "", _
</span><span>                                            x + anchoControl + 200, y, 300, altoControl)
</span><span>    With ctlCheckBox
</span><span>        .Name = "chkActivo"
</span><span>    End With
</span><span>  
</span><span>    Set ctlLabel = objAccess.CreateControl(nombreFormulario, acLabel, acDetail, ctlCheckBox.Name, "", _
</span><span>                                         x + anchoControl + 600, y, anchoControl, altoControl)
</span><span>    ctlLabel.Caption = "Registro Activo"
</span><span>  
</span><span>    y = y + altoControl + 500 ' Espacio extra antes de botones
</span><span>  
</span><span>    ' 7. CREAR BOTONES DE ACCIÓN
</span><span>    ' Botón Guardar
</span><span>    Set ctlCommandButton = objAccess.CreateControl(nombreFormulario, acCommandButton, acDetail, "", "", _
</span><span>                                                 x, y, 1200, 400)
</span><span>    With ctlCommandButton
</span><span>        .Name = "btnGuardar"
</span><span>        .Caption = "Guardar"
</span><span>        .FontBold = True
</span><span>        .BackColor = RGB(0, 128, 0) ' Verde
</span><span>        .ForeColor = RGB(255, 255, 255) ' Texto blanco
</span><span>    End With
</span><span>  
</span><span>    ' Botón Cancelar
</span><span>    Set ctlCommandButton = objAccess.CreateControl(nombreFormulario, acCommandButton, acDetail, "", "", _
</span><span>                                                 x + 1400, y, 1200, 400)
</span><span>    With ctlCommandButton
</span><span>        .Name = "btnCancelar"
</span><span>        .Caption = "Cancelar"
</span><span>        .FontBold = True
</span><span>        .BackColor = RGB(128, 0, 0) ' Rojo
</span><span>        .ForeColor = RGB(255, 255, 255) ' Texto blanco
</span><span>    End With
</span><span>  
</span><span>    ' Botón Cerrar
</span><span>    Set ctlCommandButton = objAccess.CreateControl(nombreFormulario, acCommandButton, acDetail, "", "", _
</span><span>                                                 x + 2800, y, 1200, 400)
</span><span>    With ctlCommandButton
</span><span>        .Name = "btnCerrar"
</span><span>        .Caption = "Cerrar"
</span><span>        .FontBold = True
</span><span>        .BackColor = RGB(128, 128, 128) ' Gris
</span><span>        .ForeColor = RGB(255, 255, 255) ' Texto blanco
</span><span>    End With
</span><span>  
</span><span>    ' 8. AGREGAR CÓDIGO VBA A LOS BOTONES
</span><span>    Call AgregarCodigoEventos(objAccess, nombreFormulario)
</span><span>  
</span><span>    ' 9. GUARDAR EL FORMULARIO CON NOMBRE PERSONALIZADO
</span><span>    Call GuardarFormulario(objAccess, nombreFormulario)
</span><span>  
</span><span>    Debug.Print "Formulario con controles creado exitosamente"
</span><span>    Exit Sub
</span><span>  
</span><span>ErrorHandler:
</span><span>    Debug.Print "Error al crear formulario: " & Err.Description
</span><span>End Sub
</span><span></span></code></span></div></div></div></pre>

## Función para Agregar Código a Eventos

<pre class="not-prose w-full rounded font-mono text-sm font-extralight"><div class="codeWrapper text-light selection:text-super selection:bg-super/10 my-md relative flex flex-col rounded font-mono text-sm font-normal bg-subtler"><div class="translate-y-xs -translate-x-xs bottom-xl mb-xl sticky top-0 flex h-0 items-start justify-end"></div><div class="-mt-xl"><div><div data-testid="code-language-indicator" class="text-quiet bg-subtle py-xs px-sm inline-block rounded-br rounded-tl-[3px] font-thin">text</div></div><div class="pr-lg"><span><code><span><span>Private Sub AgregarCodigoEventos(objAccess As Access.Application, nombreFormulario As String)
</span></span><span>    ' Agregar código VBA a los eventos de los botones
</span><span>    On Error Resume Next
</span><span>  
</span><span>    Dim moduloFormulario As Module
</span><span>    Dim codigoVBA As String
</span><span>  
</span><span>    ' Obtener el módulo del formulario
</span><span>    Set moduloFormulario = objAccess.Forms(nombreFormulario).Module
</span><span>  
</span><span>    ' Código para el botón Guardar
</span><span>    codigoVBA = "Private Sub btnGuardar_Click()" & vbCrLf & _
</span><span>                "    MsgBox ""Datos guardados correctamente"", vbInformation, ""Guardar""" & vbCrLf & _
</span><span>                "End Sub" & vbCrLf & vbCrLf
</span><span>  
</span><span>    moduloFormulario.InsertText codigoVBA
</span><span>  
</span><span>    ' Código para el botón Cancelar
</span><span>    codigoVBA = "Private Sub btnCancelar_Click()" & vbCrLf & _
</span><span>                "    If MsgBox(""¿Desea cancelar los cambios?"", vbYesNo + vbQuestion, ""Cancelar"") = vbYes Then" & vbCrLf & _
</span><span>                "        DoCmd.Close acForm, Me.Name, acSaveNo" & vbCrLf & _
</span><span>                "    End If" & vbCrLf & _
</span><span>                "End Sub" & vbCrLf & vbCrLf
</span><span>  
</span><span>    moduloFormulario.InsertText codigoVBA
</span><span>  
</span><span>    ' Código para el botón Cerrar
</span><span>    codigoVBA = "Private Sub btnCerrar_Click()" & vbCrLf & _
</span><span>                "    DoCmd.Close acForm, Me.Name, acSaveYes" & vbCrLf & _
</span><span>                "End Sub" & vbCrLf & vbCrLf
</span><span>  
</span><span>    moduloFormulario.InsertText codigoVBA
</span><span>  
</span><span>    ' Código para el evento Load del formulario
</span><span>    codigoVBA = "Private Sub Form_Load()" & vbCrLf & _
</span><span>                "    Me.txtFechaRegistro = Date" & vbCrLf & _
</span><span>                "    Me.chkActivo = True" & vbCrLf & _
</span><span>                "    Me.cboCategoria.ListIndex = 0" & vbCrLf & _
</span><span>                "End Sub" & vbCrLf
</span><span>  
</span><span>    moduloFormulario.InsertText codigoVBA
</span><span>  
</span><span>    Debug.Print "Código VBA agregado a los eventos"
</span><span>    On Error GoTo 0
</span><span>End Sub
</span><span></span></code></span></div></div></div></pre>

## Función para Guardar el Formulario

<pre class="not-prose w-full rounded font-mono text-sm font-extralight"><div class="codeWrapper text-light selection:text-super selection:bg-super/10 my-md relative flex flex-col rounded font-mono text-sm font-normal bg-subtler"><div class="translate-y-xs -translate-x-xs bottom-xl mb-xl sticky top-0 flex h-0 items-start justify-end"></div><div class="-mt-xl"><div><div data-testid="code-language-indicator" class="text-quiet bg-subtle py-xs px-sm inline-block rounded-br rounded-tl-[3px] font-thin">text</div></div><div class="pr-lg"><span><code><span><span>Private Sub GuardarFormulario(objAccess As Access.Application, nombreFormulario As String)
</span></span><span>    ' Guardar el formulario con un nombre personalizado
</span><span>    On Error GoTo ErrorHandler
</span><span>  
</span><span>    Dim nombreFinal As String
</span><span>    nombreFinal = "frmEjemploProgramatico"
</span><span>  
</span><span>    ' Verificar si ya existe un formulario con ese nombre
</span><span>    If FormularioExiste(objAccess, nombreFinal) Then
</span><span>        objAccess.DoCmd.DeleteObject acForm, nombreFinal
</span><span>        Debug.Print "Formulario existente eliminado: " & nombreFinal
</span><span>    End If
</span><span>  
</span><span>    ' Guardar el formulario con el nombre temporal
</span><span>    objAccess.DoCmd.Save acForm, nombreFormulario
</span><span>  
</span><span>    ' Cerrar el formulario antes de renombrar
</span><span>    objAccess.DoCmd.Close acForm, nombreFormulario, acSaveYes
</span><span>    objAccess.DoCmd.Restore
</span><span>  
</span><span>    ' Renombrar al nombre final
</span><span>    objAccess.DoCmd.Rename nombreFinal, acForm, nombreFormulario
</span><span>  
</span><span>    Debug.Print "Formulario guardado como: " & nombreFinal
</span><span>    Exit Sub
</span><span>  
</span><span>ErrorHandler:
</span><span>    Debug.Print "Error al guardar formulario: " & Err.Description
</span><span>End Sub
</span><span>
</span><span>Private Function FormularioExiste(objAccess As Access.Application, nombreFormulario As String) As Boolean
</span><span>    ' Verificar si existe un formulario con el nombre especificado
</span><span>    On Error Resume Next
</span><span>  
</span><span>    Dim obj As AccessObject
</span><span>  
</span><span>    For Each obj In objAccess.CurrentProject.AllForms
</span><span>        If obj.Name = nombreFormulario Then
</span><span>            FormularioExiste = True
</span><span>            Exit Function
</span><span>        End If
</span><span>    Next obj
</span><span>  
</span><span>    FormularioExiste = False
</span><span>    On Error GoTo 0
</span><span>End Function
</span><span></span></code></span></div></div></div></pre>

## Función de Cierre Completo

<pre class="not-prose w-full rounded font-mono text-sm font-extralight"><div class="codeWrapper text-light selection:text-super selection:bg-super/10 my-md relative flex flex-col rounded font-mono text-sm font-normal bg-subtler"><div class="translate-y-xs -translate-x-xs bottom-xl mb-xl sticky top-0 flex h-0 items-start justify-end"></div><div class="-mt-xl"><div><div data-testid="code-language-indicator" class="text-quiet bg-subtle py-xs px-sm inline-block rounded-br rounded-tl-[3px] font-thin">text</div></div><div class="pr-lg"><span><code><span><span>Private Sub CerrarAccessCompleto(objAccess As Access.Application)
</span></span><span>    ' Cerrar Access de forma completa y desatendida
</span><span>    On Error Resume Next
</span><span>  
</span><span>    ' 1. Cerrar todos los formularios abiertos
</span><span>    Do While objAccess.Forms.Count > 0
</span><span>        objAccess.DoCmd.Close acForm, objAccess.Forms(0).Name, acSaveYes
</span><span>    Loop
</span><span>  
</span><span>    ' 2. Cerrar todos los informes abiertos
</span><span>    Do While objAccess.Reports.Count > 0
</span><span>        objAccess.DoCmd.Close acReport, objAccess.Reports(0).Name, acSaveNo
</span><span>    Loop
</span><span>  
</span><span>    ' 3. Cerrar base de datos actual
</span><span>    objAccess.CloseCurrentDatabase
</span><span>  
</span><span>    ' 4. Ejecutar Quit
</span><span>    objAccess.Quit acQuitSaveAll
</span><span>  
</span><span>    ' 5. Liberar objeto
</span><span>    Set objAccess = Nothing
</span><span>  
</span><span>    ' 6. Forzar recolección de basura
</span><span>    DoEvents
</span><span>    DoEvents
</span><span>  
</span><span>    Debug.Print "Access cerrado completamente"
</span><span>    On Error GoTo 0
</span><span>End Sub
</span><span></span></code></span></div></div></div></pre>

## Ejemplo de Uso Alternativo con Tabla de Datos

<pre class="not-prose w-full rounded font-mono text-sm font-extralight"><div class="codeWrapper text-light selection:text-super selection:bg-super/10 my-md relative flex flex-col rounded font-mono text-sm font-normal bg-subtler"><div class="translate-y-xs -translate-x-xs bottom-xl mb-xl sticky top-0 flex h-0 items-start justify-end"></div><div class="-mt-xl"><div><div data-testid="code-language-indicator" class="text-quiet bg-subtle py-xs px-sm inline-block rounded-br rounded-tl-[3px] font-thin">text</div></div><div class="pr-lg"><span><code><span><span>Public Sub CrearFormularioConDatos()
</span></span><span>    ' Crear formulario conectado a una tabla de datos
</span><span>  
</span><span>    Dim objAccess As Access.Application
</span><span>    Dim rutaBD As String
</span><span>  
</span><span>    rutaBD = "C:\MiBaseDatos\Ejemplo.accdb"
</span><span>  
</span><span>    Set objAccess = AbrirAccessSilencioso(rutaBD)
</span><span>  
</span><span>    If Not objAccess Is Nothing Then
</span><span>        Call CrearFormularioConTabla(objAccess, "Clientes")
</span><span>        Call CerrarAccessCompleto(objAccess)
</span><span>    End If
</span><span>  
</span><span>    Set objAccess = Nothing
</span><span>End Sub
</span><span>
</span><span>Private Sub CrearFormularioConTabla(objAccess As Access.Application, nombreTabla As String)
</span><span>    ' Crear formulario basado en una tabla existente
</span><span>    On Error GoTo ErrorHandler
</span><span>  
</span><span>    Dim frm As Access.Form
</span><span>    Dim rst As DAO.Recordset
</span><span>    Dim fld As DAO.Field
</span><span>    Dim x As Long, y As Long
</span><span>    Dim ctlLabel As Access.Control
</span><span>    Dim ctlTextBox As Access.Control
</span><span>  
</span><span>    ' Verificar que la tabla existe
</span><span>    Set rst = objAccess.CurrentDb.OpenRecordset(nombreTabla)
</span><span>  
</span><span>    Set frm = objAccess.CreateForm()
</span><span>    frm.RecordSource = nombreTabla
</span><span>    frm.Caption = "Formulario de " & nombreTabla
</span><span>  
</span><span>    x = 500
</span><span>    y = 500
</span><span>  
</span><span>    ' Crear controles para cada campo de la tabla
</span><span>    For Each fld In rst.Fields
</span><span>        ' Crear etiqueta
</span><span>        Set ctlLabel = objAccess.CreateControl(frm.Name, acLabel, acDetail, "", "", _
</span><span>                                             x, y, 1500, 300)
</span><span>        ctlLabel.Caption = fld.Name & ":"
</span><span>      
</span><span>        ' Crear control de texto vinculado al campo
</span><span>        Set ctlTextBox = objAccess.CreateControl(frm.Name, acTextBox, acDetail, "", fld.Name, _
</span><span>                                               x + 1700, y, 2500, 300)
</span><span>        ctlTextBox.Name = "txt" & fld.Name
</span><span>      
</span><span>        y = y + 500 ' Siguiente línea
</span><span>    Next fld
</span><span>  
</span><span>    rst.Close
</span><span>    Set rst = Nothing
</span><span>  
</span><span>    ' Guardar formulario
</span><span>    objAccess.DoCmd.Save acForm, frm.Name
</span><span>    objAccess.DoCmd.Close acForm, frm.Name, acSaveYes
</span><span>  
</span><span>    Debug.Print "Formulario basado en tabla creado: " & frm.Name
</span><span>    Exit Sub
</span><span>  
</span><span>ErrorHandler:
</span><span>    Debug.Print "Error creando formulario con tabla: " & Err.Description
</span><span>    If Not rst Is Nothing Then rst.Close
</span><span>    Set rst = Nothing
</span><span>End Sub
</span><span></span></code></span></div></div></div></pre>

## Consideraciones Importantes

## Propiedades Críticas para Controles[](https://learn.microsoft.com/en-us/office/vba/api/access.application.createcontrol)

1. **CreateControl solo funciona en Vista de Diseño**[](https://stackoverflow.com/questions/31301070/how-to-create-controls-at-run-time-access-vb)
2. **Los controles deben tener nombres únicos**
3. **Las coordenadas están en twips (1440 twips = 1 pulgada)**
4. **El formulario debe estar abierto en modo de diseño durante la creación**[](https://www.pcreview.co.uk/threads/dynamically-add-control-to-msacess-form.2383165/)

## Constantes de Tipos de Control Disponibles[](https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2003/aa221167(v=office.11))

* `acTextBox` - Cuadro de texto
* `acLabel` - Etiqueta
* `acCommandButton` - Botón de comando
* `acComboBox` - Cuadro combinado
* `acListBox` - Cuadro de lista
* `acCheckBox` - Casilla de verificación
* `acOptionButton` - Botón de opción
* `acToggleButton` - Botón de alternancia
* `acImage` - Imagen
* `acRectangle` - Rectángulo
* `acLine` - Línea
* `acSubform` - Subformulario
* `acTabCtl` - Control de pestañas

Este ejemplo completo demuestra la creación totalmente programática y desatendida de un formulario con múltiples tipos de controles, incluyendo eventos VBA, siguiendo la documentación oficial de Microsoft Access.[](https://learn.microsoft.com/en-us/office/vba/api/access.application.createform)





' Script VBS: Creación programática de formulario en Access totalmente desatendido

Option Explicit
On Error Resume Next

Dim objAccess, rutaBD
rutaBD = "C:\MiBaseDatos\Ejemplo.accdb"

' 1. Abrir Access silencioso
Set objAccess = CreateObject("Access.Application")
objAccess.Visible = False
objAccess.UserControl = False

objAccess.SetOption "Confirm Record Changes", False
objAccess.SetOption "Confirm Document Deletions", False
objAccess.SetOption "Confirm Action Queries", False
objAccess.SetOption "Show Status Bar", False

objAccess.OpenCurrentDatabase rutaBD, False

' 2. Crear formulario y controles
Call CrearFormularioConControles(objAccess)

' 3. Cerrar Access completamente
Call CerrarAccessCompleto(objAccess)

Set objAccess = Nothing
WScript.Quit(0)

' ================================
Sub CrearFormularioConControles(a)
    Const acDetail = 0
    Const acHeader = 1
    Const acLabel = 1004
    Const acTextBox = 109
    Const acComboBox = 111
    Const acCheckBox = 106
    Const acCommandButton = 104

    Dim frmName, ctl, x, y, w, h

    ' Crear formulario
    frmName = a.CreateForm().Name

    ' Propiedades del formulario
    a.Forms(frmName).Caption           = "Formulario VBS"
    a.Forms(frmName).AutoCenter        = True
    a.Forms(frmName).PopUp             = False
    a.Forms(frmName).Modal             = False
    a.Forms(frmName).Section(acHeader).Visible = True
    a.Forms(frmName).Section(acHeader).Height  = 800
    a.Forms(frmName).Section(acDetail).Height  = 6000

    x = 500: y = 200: w = 2000: h = 300

    ' Etiqueta título
    Set ctl = a.CreateControl(frmName, acLabel, acHeader, , , 2000, 100, 4000, 400)
    ctl.Caption   = "FORMULARIO VBS"
    ctl.FontBold  = True
    ctl.FontSize  = 14

    ' Campo Nombre
    Set ctl = a.CreateControl(frmName, acLabel, acDetail, , , x, y, w, h)
    ctl.Caption = "Nombre:"
    Set ctl = a.CreateControl(frmName, acTextBox, acDetail, , , x + w + 200, y, w + 500, h)
    ctl.Name    = "txtNombre"
    y = y + h + 300

    ' Campo Email
    Set ctl = a.CreateControl(frmName, acLabel, acDetail, , , x, y, w, h)
    ctl.Caption = "Email:"
    Set ctl = a.CreateControl(frmName, acTextBox, acDetail, , , x + w + 200, y, w + 500, h)
    ctl.Name    = "txtEmail"
    y = y + h + 300

    ' Combo Categoría
    Set ctl = a.CreateControl(frmName, acLabel, acDetail, , , x, y, w, h)
    ctl.Caption = "Categoría:"
    Set ctl = a.CreateControl(frmName, acComboBox, acDetail, , , x + w + 200, y, w + 500, h)
    ctl.RowSourceType = "Value List"
    ctl.RowSource     = "Cliente;Proveedor;Empleado;Otro"
    ctl.LimitToList   = True
    y = y + h + 300

    ' Fecha Registro
    Set ctl = a.CreateControl(frmName, acLabel, acDetail, , , x, y, w, h)
    ctl.Caption = "Fecha Registro:"
    Set ctl = a.CreateControl(frmName, acTextBox, acDetail, , , x + w + 200, y, w, h)
    ctl.Format       = "Short Date"
    ctl.ShowDatePicker = 1
    y = y + h + 300

    ' Checkbox Activo
    Set ctl = a.CreateControl(frmName, acCheckBox, acDetail, , , x + w + 200, y, 300, h)
    ctl.Name = "chkActivo"
    Set ctl = a.CreateControl(frmName, acLabel, acDetail, ctl.Name, , x + w + 600, y, w, h)
    ctl.Caption = "Registro Activo"
    y = y + h + 500

    ' Botones
    Set ctl = a.CreateControl(frmName, acCommandButton, acDetail, , , x, y, 1200, 400)
    ctl.Name    = "btnGuardar": ctl.Caption = "Guardar"
    Set ctl = a.CreateControl(frmName, acCommandButton, acDetail, , , x + 1400, y, 1200, 400)
    ctl.Name    = "btnCancelar": ctl.Caption = "Cancelar"
    Set ctl = a.CreateControl(frmName, acCommandButton, acDetail, , , x + 2800, y, 1200, 400)
    ctl.Name    = "btnCerrar": ctl.Caption = "Cerrar"

    ' Agregar código a eventos
    Call AgregarCodigoEventosVBS(a, frmName)

    ' Guardar y renombrar formulario
    Call GuardarYRenombrar(a, frmName, "frmEjemploVBS")
End Sub

Sub AgregarCodigoEventosVBS(a, frmName)
    Dim modCode
    modCode  = "Private Sub btnGuardar_Click()" & vbCrLf
    modCode &= "    MsgBox ""Guardado exitoso"", vbInformation" & vbCrLf
    modCode &= "End Sub" & vbCrLf & vbCrLf
    modCode &= "Private Sub btnCancelar_Click()" & vbCrLf
    modCode &= "    DoCmd.Close acForm, Me.Name, acSaveNo" & vbCrLf
    modCode &= "End Sub" & vbCrLf & vbCrLf
    modCode &= "Private Sub btnCerrar_Click()" & vbCrLf
    modCode &= "    DoCmd.Close acForm, Me.Name, acSaveYes" & vbCrLf
    modCode &= "End Sub" & vbCrLf & vbCrLf
    modCode &= "Private Sub Form_Load()" & vbCrLf
    modCode &= "    Me.txtFechaRegistro = Date" & vbCrLf
    modCode &= "    Me.chkActivo = True" & vbCrLf
    modCode &= "    Me.cboCategoria.ListIndex = 0" & vbCrLf
    modCode &= "End Sub"
    a.Modules(frmName).InsertText modCode
End Sub

Sub GuardarYRenombrar(a, tempName, finalName)
    Const acForm = 2, acSaveYes = True
    If a.CurrentProject.AllForms(finalName).Name <> "" Then
        a.DoCmd.DeleteObject acForm, finalName
    End If
    a.DoCmd.Save acForm, tempName
    a.DoCmd.Close acForm, tempName, acSaveYes
    a.DoCmd.Rename finalName, acForm, tempName
End Sub

Sub CerrarAccessCompleto(a)
    Const acForm = 2, acReport = 3, acSaveYes = True, acQuitSaveAll = 0
    Do While a.Forms.Count > 0
        a.DoCmd.Close acForm, a.Forms(0).Name, acSaveYes
    Loop
    Do While a.Reports.Count > 0
        a.DoCmd.Close acReport, a.Reports(0).Name, False
    Loop
    a.CloseCurrentDatabase
    a.Quit acQuitSaveAll
End Sub

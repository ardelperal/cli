# CLI para Microsoft Access

Herramienta de línea de comandos desarrollada en VBScript para gestionar bases de datos de Microsoft Access y sus objetos. Diseñada para integración continua con proyectos de Access y VBA.

## Características Principales

- **Gestión de objetos de Access**: Listar tablas, formularios, consultas y módulos
- **Exportación de formularios**: Exportar formularios de Access a formato JSON
- **Soporte para bases de datos protegidas**: Manejo de contraseñas
- **Configuración flexible**: Archivo INI para parámetros de configuración
- **Modo de prueba**: Simulación de operaciones con `--dry-run`
- **Logging detallado**: Registro de operaciones con modo verbose

## Comandos Disponibles

### list-objects
Lista todos los objetos de una base de datos de Access.

```bash
cscript cli.vbs list-objects <db_path> [--password <pwd>] [--schema] [--output]
```

**Parámetros:**
- `<db_path>`: Ruta a la base de datos de Access
- `--password <pwd>`: Contraseña de la base de datos (opcional)
- `--schema`: Muestra detalles de campos en las tablas (opcional)
- `--output`: Exporta resultados a archivo `[nombre_bd]_listobjects.txt` (opcional)

**Ejemplo:**
```bash
cscript cli.vbs list-objects Expedientes.accdb --password mipassword --schema --output
```

### export-form
Exporta un formulario de Access a formato JSON.

```bash
cscript cli.vbs export-form <db_path> <form_name> [--password <pwd>]
```

**Parámetros:**
- `<db_path>`: Ruta a la base de datos de Access
- `<form_name>`: Nombre del formulario a exportar
- `--password <pwd>`: Contraseña de la base de datos (opcional)

**Ejemplo:**
```bash
cscript cli.vbs export-form Expedientes.accdb FormularioTest --password mipassword
```

### Opciones Globales

- `--dry-run`: Simula la operación sin ejecutarla
- `--verbose`: Muestra información detallada de la operación
- `--help`: Muestra la ayuda del comando

## Configuración

El archivo `cli.ini` contiene la configuración de la herramienta:

```ini
[GENERAL]
DryRun = false
Verbose = false
LogFile = cli.log

[ACCESS]
AutomationSecurity = 3
DefaultOpenMode = 1
DefaultRecordLocking = 0

[UI]
Root = .\ui
FormsDir = forms
AssetsDir = assets
AssetsImgDir = img
AssetsImgExtensions = .png,.jpg,.jpeg,.gif,.bmp
IncludeSubdirectories = true
FormFilePattern = *.json
NameFromFileBase = true
```

## Estructura de Archivos

```
cli/
├── cli.vbs              # Script principal
├── cli.ini              # Archivo de configuración
├── cli.log              # Archivo de log
├── README.md            # Este archivo
├── docs/                # Documentación
├── ui/                  # Archivos de interfaz
│   └── forms/           # Formularios exportados en JSON
└── assets/              # Recursos adicionales
```

## Requisitos

- Microsoft Access instalado
- Windows Script Host (WSH)
- Permisos para acceder al modelo de objetos VBA de Access

## Uso Básico

1. **Listar objetos de una base de datos:**
   ```bash
   cscript cli.vbs list-objects MiBaseDatos.accdb
   ```

2. **Exportar un formulario:**
   ```bash
   cscript cli.vbs export-form MiBaseDatos.accdb MiFormulario
   ```

3. **Modo de prueba:**
   ```bash
   cscript cli.vbs --dry-run list-objects MiBaseDatos.accdb
   ```

4. **Con información detallada:**
   ```bash
   cscript cli.vbs --verbose list-objects MiBaseDatos.accdb --schema
   ```

## Notas Técnicas

- La herramienta maneja automáticamente la apertura y cierre seguro de Access
- Soporta bases de datos con y sin contraseña
- Los formularios se exportan con toda su estructura de controles y propiedades
- El modo `--schema` incluye información detallada de tipos de campos en tablas
- Los logs se guardan automáticamente en `cli.log`

## Estado del Desarrollo

✅ **Completado:**
- Comando `list-objects` con soporte completo para parámetros
- Comando `export-form` funcional
- Manejo seguro de bases de datos con contraseña
- Sistema de configuración INI
- Logging y modo verbose

🔄 **En desarrollo:**
- Comandos adicionales para gestión de módulos VBA
- Importación de formularios desde JSON
- Sincronización bidireccional - Estado Final

Herramienta de desarrollo implementada en VBScript para gestionar bases de datos de Microsoft Access desde línea de comandos.

## Características Implementadas

### Funcionalidades Core
- **Extracción de módulos VBA**: Exporta módulos y clases a archivos fuente (.bas, .cls)
- **Reconstrucción de módulos**: Importa módulos desde archivos fuente con validación de sintaxis
- **Gestión de formularios**: Exportación e importación completa de formularios con propiedades, controles y código VBA
- **Esquemas de base de datos**: Exportación de estructura de tablas en formato JSON y Markdown
- **Listado de objetos**: Inventario completo de todos los objetos de la base de datos

### Mejoras de Robustez Implementadas
- **Normalización de extensiones**: Uso de `ResolveSourcePathForModule` para manejo consistente de archivos
- **Compilación centralizada**: Una sola compilación al final de las operaciones para mayor eficiencia
- **Función AntiUI**: Centralización de la configuración anti-interactividad en una función reutilizable
- **Sintaxis corregida**: Eliminación de `app.Echo = False` (sintaxis incorrecta) y `app.DisplayAlerts = False` (no existe en Access)
- **Reintentos robustos**: Lógica de reintentos para obtener nombres de formularios activos tras `CreateForm`
- **Limpieza automática**: Manejo de errores con limpieza de recursos en caso de fallos
- **Verificación de referencias**: `CheckMissingReferences` detecta referencias VBA rotas antes de compilar
- **Acceso VBE obligatorio**: `CheckVBProjectAccess` es obligatorio en operaciones críticas
- **Operaciones DryRun**: Funciones `Maybe_DoCmd_*` permiten simulación sin efectos reales
- **Manejo de errores estructurado**: `HandleErr` y `HandleErrWithCleanup` para gestión centralizada de errores

### Configuración
- **Archivo INI**: Configuración centralizada en `cli.ini`
- **Validación**: Comando `/validate` para verificar configuración
- **Modo simulación**: Modificador `/dry-run` para pruebas sin cambios reales
- **Verificación VBE**: Validación obligatoria del acceso al modelo de objetos VBA
- **Detección de referencias**: Verificación automática de referencias rotas antes de compilar

### Testing y Depuración
- **Modificadores de testing**: `/dry-run`, `/validate`
- **Logging detallado**: Opciones `--verbose` y `--quiet`
- **Manejo de errores**: Captura y reporte de errores con contexto
- **Simulación segura**: Operaciones DryRun que no modifican la base de datos
- **Verificación previa**: Validación de referencias y acceso VBE antes de operaciones críticas

## Configuración

### Archivo cli.ini
El archivo `cli.ini` contiene toda la configuración necesaria:

```ini
[DATABASE]
DefaultPath=CONDOR.accdb
Password=

[MODULES]
SourcePath=src
Extensions=.bas,.cls
ExcludePatterns=

[VBA]
EnableVBEAccess=true
TrustVBAProject=true
RequireVBAAccess=true
```

### Requisitos de Configuración VBA
Para que funcione correctamente, debe habilitarse el acceso al modelo de objetos VBA:

1. Abrir Microsoft Access
2. Ir a **Archivo > Opciones > Centro de confianza**
3. Hacer clic en **Configuración del centro de confianza**
4. Seleccionar **Configuración de macros**
5. Marcar **"Confiar en el acceso al modelo de objetos de proyectos VBA"**

## Uso

### Comandos Disponibles

#### Exportar módulos
```cmd
cscript //NoLogo cli.vbs export-modules [/verbose]
```

#### Exportar módulo específico
```cmd
cscript //NoLogo cli.vbs export-module NombreModulo [/verbose]
```

#### Actualizar módulo
```cmd
cscript //NoLogo cli.vbs update NombreModulo [/verbose]
```

#### Listar módulos
```cmd
cscript //NoLogo cli.vbs list-modules [/verbose]
```

### Modificadores
- `/verbose` - Activa el modo detallado con información adicional de diagnóstico

## Estructura del Proyecto
```
cli/
├── cli.vbs              # Script principal
├── cli.ini              # Archivo de configuración
├── CONDOR.accdb         # Base de datos de ejemplo
├── src/                 # Módulos fuente
│   ├── CAppManager.cls  # Ejemplo de clase
│   └── ...
├── assets/
│   └── condor_cli.vbs   # Referencia de implementación
└── README.md            # Este archivo
```

## Funciones Principales

### CheckMissingReferences
Función que verifica referencias VBA rotas antes de compilar:
- Detecta referencias no disponibles en el proyecto VBA
- Previene errores de compilación por dependencias faltantes
- Se ejecuta automáticamente antes de todas las operaciones de compilación

### Maybe_DoCmd_* (Operaciones DryRun)
Conjunto de funciones que respetan el modo DryRun:
- `Maybe_DoCmd_Rename`: Renombrado seguro de objetos
- `Maybe_DoCmd_Delete`: Eliminación controlada de objetos
- `Maybe_DoCmd_OpenFormDesignHidden`: Apertura de formularios en modo diseño
- `Maybe_DoCmd_Close`: Cierre controlado de objetos
- `Maybe_DoCmd_Save`: Guardado seguro de objetos
- `Maybe_DoCmd_CompileAndSaveAllModules`: Compilación con soporte DryRun

### HandleErr y HandleErrWithCleanup
Sistema de manejo de errores estructurado:
- Diferentes niveles de severidad: "warning", "error", "critical"
- Limpieza automática de recursos en errores críticos
- Logging contextualizado según el nivel de error
- Terminación controlada de la aplicación cuando es necesario

### ImportVBAModuleSafe
Función robusta para importar módulos VBA con manejo de errores mejorado:
- Verificación de disponibilidad de Access y VBE
- Manejo seguro de errores de acceso al proyecto VBA
- Mensajes informativos sobre configuración requerida

### Manejo de Errores
El sistema incluye manejo robusto de errores para:
- Problemas de acceso al VBE (Visual Basic Editor)
- Configuración de confianza VBA
- Archivos no encontrados
- Módulos inexistentes

## Testing
El CLI es testeable usando modificadores directamente:
```cmd
cscript //NoLogo cli.vbs update CAppManager /verbose
```

## Notas Importantes
- **NUNCA** incluye caracteres especiales ni tildes en la salida de terminal
- Requiere configuración manual del acceso VBA en Access
- Los módulos se almacenan en la carpeta `src/` por defecto
- Soporta bases de datos con y sin contraseña

## Estado del Desarrollo
- ✅ Exportación de módulos funcional
- ✅ Listado de módulos funcional  
- ⚠️ Importación/actualización requiere configuración VBA manual
- ✅ Configuración mediante INI
- ✅ Logging y modo verbose
- ✅ Manejo robusto de errores

## Solución de Problemas

### Error "Proyecto VBA no está disponible"
Este error indica que el acceso al modelo de objetos VBA no está habilitado. Siga los pasos de configuración VBA mencionados anteriormente.

### Error "Se requiere un objeto - VBProject"
Similar al anterior, requiere habilitar la confianza en el acceso al modelo de objetos VBA en la configuración de Access.

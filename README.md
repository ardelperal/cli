# CLI para Microsoft Access - Estado Final

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

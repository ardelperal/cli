# CLI.VBS - Herramienta de Desarrollo para Microsoft Access

## Descripción
CLI.VBS es una herramienta de línea de comandos desarrollada en VBScript para gestionar módulos VBA en bases de datos de Microsoft Access. Permite exportar, importar y actualizar módulos de forma automatizada.

## Características Principales
- Exportación de módulos VBA desde Access a archivos fuente
- Importación de módulos desde archivos fuente a Access
- Actualización de módulos existentes
- Soporte para módulos estándar (.bas) y clases (.cls)
- Configuración mediante archivo INI
- Logging detallado de operaciones
- Modo verbose para diagnóstico

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

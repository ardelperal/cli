# CLI.vbs - Herramienta de Desarrollo para Microsoft Access

Una herramienta de línea de comandos desarrollada en VBScript para trabajar con bases de datos de Microsoft Access, con soporte completo para rutas relativas y portabilidad.

## Características Principales

- **Portabilidad completa**: Todas las rutas son relativas al directorio del CLI
- **Extracción de datos**: Exporta tablas, informes y consultas a JSON
- **Gestión de módulos VBA**: Rebuild y update de módulos desde archivos fuente
- **Sistema de testing**: Modificadores `/dryrun` y `/verbose` para pruebas
- **Configuración flexible**: Archivo INI con parámetros personalizables
- **Logging avanzado**: Sistema de registro con diferentes niveles
- **Manejo seguro de Access**: Apertura y cierre controlado de aplicaciones

## 📦 Instalación

1. Clona o descarga los archivos del proyecto
2. Asegúrate de tener Microsoft Access instalado
3. Coloca tus bases de datos Access en el mismo directorio del CLI o usa rutas relativas

## ⚙️ Configuración

El archivo `cli.ini` contiene toda la configuración. **Todas las rutas son relativas al directorio donde está el CLI**, lo que hace la herramienta completamente portable:

```ini
[DATABASE]
; Ruta relativa a la base de datos
DefaultPath=sample.accdb
Password=
Timeout=30

[OUTPUT]
; Directorio de salida relativo
DefaultPath=output
Format=json
PrettyPrint=true

[LOGGING]
; Archivo de log relativo
LogFile=cli.log
LogLevel=INFO
Verbose=false

[EXTRACTION]
IncludeTables=true
IncludeForms=true
IncludeQueries=true
IncludeRelations=true
FilterSystemObjects=true
```

## Uso

### Comandos Principales

#### Extraer toda la información
```cmd
cscript cli.vbs extract-all "C:\ruta\basedatos.accdb" "C:\salida\datos.json"
```

#### Extraer solo tablas
```cmd
cscript cli.vbs extract-tables "C:\ruta\basedatos.accdb" "C:\salida\tablas.json"
```

#### Extraer solo informes

```cmd
cscript cli.vbs extract-reports "C:\ruta\basedatos.accdb" "C:\salida\informes.json"
```

#### Listar objetos de la base de datos
```cmd
cscript cli.vbs list-objects "C:\ruta\basedatos.accdb"
```

### Modificadores de Testing

#### Modo dry-run (simulación)
```cmd
cscript cli.vbs extract-all "C:\mi_base.accdb" /dry-run
```

#### Modo verbose (información detallada)
```cmd
cscript cli.vbs extract-all "C:\ruta\bd.accdb" "C:\salida.json" /verbose
```

#### Modo debug (información de depuración)
```cmd
cscript cli.vbs extract-all "C:\ruta\bd.accdb" "C:\salida.json" /debug
```

#### Modo dry-run (simulación)
```cmd
cscript cli.vbs extract-all "C:\ruta\bd.accdb" "C:\salida.json" /dryrun
```

### Ayuda
```cmd
cscript cli.vbs help
```

## Estructura de Salida JSON

### Información de Tablas
```json
{
  "tables": {
    "NombreTabla": {
      "name": "NombreTabla",
      "recordCount": 150,
      "dateCreated": "2024-01-15 10:30:00",
      "lastUpdated": "2024-01-20 14:45:00",
      "fields": [
        {
          "name": "ID",
          "type": "Long",
          "size": 4,
          "required": true,
          "allowZeroLength": false
        }
      ],
      "indexes": [
        {
          "name": "PrimaryKey",
          "primary": true,
          "unique": true,
          "fields": ["ID"]
        }
      ]
    }
  }
}
```

### Información de Informes
```json
{
  "NombreInforme": {
    "name": "NombreInforme",
    "type": "Report",
    "caption": "Título del Informe",
    "recordSource": "TablaOrigen",
    "properties": {
      "width": 8000,
      "height": 6000
    }
  }
}
```

## Funcionalidades Implementadas

### Extracción de Tablas
- ✅ Información básica (nombre, fechas, conteo de registros)
- ✅ Campos con tipos, tamaños y propiedades
- ✅ Índices primarios y secundarios
- ✅ Reglas de validación

### Extracción de Relaciones
- ✅ Relaciones entre tablas
- ✅ Campos relacionados
- ✅ Atributos de integridad referencial

### Extracción de Informes
- ✅ Propiedades del informe
- ✅ Controles y sus propiedades
- ✅ Posicionamiento y formato
- ✅ Origen de datos y enlaces

### Sistema de Testing
- ✅ Test de configuración
- ✅ Test de funciones utilitarias
- ✅ Test de conexión a base de datos
- ✅ Modificadores de línea de comandos

### Utilidades
- ✅ Conversión de tipos de datos Access
- ✅ Conversión de tipos de controles
- ✅ Sistema de logging configurable
- ✅ Manejo seguro de errores

## Requisitos del Sistema

- Windows con VBScript habilitado
- Microsoft Access instalado (para acceso a objetos COM)
- Permisos de lectura en las bases de datos objetivo
- Permisos de escritura en directorios de salida

## Ejemplos de Uso Práctico

### Documentar estructura de base de datos
```cmd
cscript cli.vbs extract-all "C:\Proyecto\datos.accdb" "C:\Docs\estructura.json" /verbose
```

### Verificar configuración antes de ejecutar
```cmd
cscript cli.vbs extract-all "C:\Proyecto\datos.accdb" "C:\Docs\estructura.json" /dryrun
```

### Ejecutar con información detallada
```cmd
cscript cli.vbs extract-all "C:\mi_base.accdb" /verbose
```

## Solución de Problemas

### Error de conexión a Access
- Verificar que Microsoft Access esté instalado
- Comprobar permisos de lectura en el archivo de base de datos
- Verificar que la ruta del archivo sea correcta

### Error de escritura de archivos
- Verificar permisos de escritura en el directorio de salida
- Comprobar que el directorio de salida exista
- Verificar espacio disponible en disco

### Tests fallando
- Ejecutar con `/debug` para información detallada
- Verificar configuración en `cli.ini`
- Comprobar que las rutas configuradas sean válidas

## Desarrollo y Contribución

El proyecto está estructurado de forma modular:

- **Configuración**: Carga desde `cli.ini`
- **Conexión Access**: Manejo seguro de objetos COM
- **Extracción**: Funciones especializadas por tipo de objeto
- **Salida**: Generación de JSON estructurado
- **Testing**: Sistema integrado de pruebas
- **Logging**: Sistema configurable de logs

Basado en la estructura y funcionalidades de `condor_cli.vbs` pero adaptado para extracción completa de información de bases de datos Access.

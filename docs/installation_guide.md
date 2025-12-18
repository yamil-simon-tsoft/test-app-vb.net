# Guía de Instalación y Configuración

## Prerrequisitos del Sistema

### Software Requerido

1. **Sistema Operativo**: Windows 10/11 (versión 1903 o superior)
2. **Python**: Versión 3.8 o superior
3. **Microsoft Word**: Office 2016 o superior
4. **Windows Application Driver (WinAppDriver)**
5. **Modo de Desarrollador habilitado**

### Verificar Prerrequisitos

#### 1. Verificar Python
```powershell
python --version
```
Debe mostrar Python 3.8+

#### 2. Verificar Word
Abrir Word manualmente para confirmar que funciona correctamente.

#### 3. Verificar Modo de Desarrollador
- Ir a Configuración > Actualización y seguridad > Para desarrolladores
- Activar "Modo de desarrollador"

## Instalación Paso a Paso

### 1. Clonar o Descargar el Proyecto
```powershell
# Si tienes Git
git clone <url-del-repositorio>
cd nombre-del-proyecto

# O descargar ZIP y extraer
```

### 2. Ejecutar Script de Activación
```powershell
# Ejecutar desde la raíz del proyecto
.\\activar.ps1
```

Este script automáticamente:
- Verifica Python
- Crea entorno virtual
- Instala dependencias
- Descarga WinAppDriver
- Verifica configuración

### 3. Configuración Manual (si el script falla)

#### Crear Entorno Virtual
```powershell
python -m venv venv
venv\\Scripts\\Activate.ps1
```

#### Instalar Dependencias
```powershell
pip install -r requirements.txt
```

#### Descargar WinAppDriver
1. Ir a https://github.com/Microsoft/WinAppDriver/releases
2. Descargar la última versión
3. Instalar en `C:\\Program Files (x86)\\Windows Application Driver\\`

### 4. Configurar Variables de Entorno

Crear archivo `.env` en la raíz del proyecto:
```
# Configuración de Word
WORD_APP_PATH=C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE
WINAPPDRIVER_PATH=C:\\Program Files (x86)\\Windows Application Driver\\WinAppDriver.exe

# Configuración de timeouts
DEFAULT_TIMEOUT=10
ELEMENT_TIMEOUT=5

# Configuración de logs
LOG_LEVEL=INFO
SCREENSHOTS_ON_ERROR=true
```

## Verificación de Instalación

### 1. Ejecutar Validación
```powershell
python src\\utils\\config.py
```

### 2. Ejecutar el Ejemplo de Word
```powershell
python examples\\word_examples\\01_word_basic_operations.py
```

### 3. Verificar Archivos Generados
- Log del ejemplo: `reports/logs/word_basic_operations.log`
- Capturas de pantalla: `reports/screenshots/`

## Solución de Problemas Comunes

### Error: "WinAppDriver not found"
**Solución**: 
1. Verificar que WinAppDriver esté instalado
2. Comprobar la ruta en `.env`
3. Ejecutar como administrador si es necesario

### Error: "Word application not found"
**Solución**:
1. Verificar instalación de Word
2. Comprobar ruta en `.env`
3. Usar `Get-Command winword` en PowerShell para encontrar la ruta

### Error: "Developer mode not enabled"
**Solución**:
1. Activar modo de desarrollador en configuración de Windows
2. Reiniciar después de activar

### Error: "Permission denied"
**Solución**:
1. Ejecutar PowerShell como administrador
2. Verificar política de ejecución: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

### Error: "Module not found"
**Solución**:
1. Activar entorno virtual: `venv\\Scripts\\Activate.ps1`
2. Reinstalar dependencias: `pip install -r requirements.txt`

## Configuración Avanzada

### Personalizar Timeouts
En `.env`:
```
DEFAULT_TIMEOUT=15
ELEMENT_TIMEOUT=8
```

### Configurar Logging
```
LOG_LEVEL=DEBUG  # Para más detalle
SCREENSHOTS_ON_ERROR=false  # Para deshabilitar capturas
```

### Rutas Personalizadas
```
WORD_APP_PATH=C:\\Custom\\Path\\WINWORD.EXE
REPORTS_PATH=C:\\Custom\\Reports\\
```

## Configuración de VS Code

### Extensiones Recomendadas
El archivo `.vscode/extensions.json` incluye:
- Python
- Pylance
- PowerShell
- Autodocstring

### Uso Básico
1. Abrir proyecto: `code .`
2. Terminal integrado: `Ctrl+`` para ejecutar comandos Python
3. Ejecutar el ejemplo directamente desde el terminal

## Mantenimiento

### Actualizar Dependencias
```powershell
pip install --upgrade -r requirements.txt
```

### Limpiar Reportes
```powershell
# Limpiar manualmente:
Remove-Item -Path "reports\\*" -Recurse -Force
```

### Verificar Funcionamiento
```powershell
python examples\\word_examples\\01_word_basic_operations.py
```

## Próximos Pasos

Después de la instalación exitosa:
1. Ejecutar el ejemplo: `python examples/word_examples/01_word_basic_operations.py`
2. Revisar el log generado en `reports/logs/`
3. Analizar las capturas de pantalla en `reports/screenshots/`
4. Personalizar configuración según necesidades
5. Extender el ejemplo con más funcionalidades de Word

Para soporte adicional, revisar:
- README.md principal
- Logs en `reports/logs/`
- Documentación específica de Word en `docs/word_automation_guide.md`
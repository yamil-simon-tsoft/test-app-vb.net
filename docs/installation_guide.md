# Guía de Instalación

## Requisitos Previos

1. **Windows 10/11**
2. **Python 3.8+** - https://python.org
3. **Microsoft Word 2016+**
4. **WinAppDriver** - https://github.com/Microsoft/WinAppDriver/releases
5. **Modo de Desarrollador** activado en Windows

## Pasos de Instalación

### 1. Instalar WinAppDriver

```powershell
# Descargar e instalar desde:
# https://github.com/Microsoft/WinAppDriver/releases

# Verificar instalación
Get-Command WinAppDriver
```

### 2. Activar Modo de Desarrollador

1. Ir a **Configuración** > **Actualización y seguridad** > **Para desarrolladores**
2. Activar **Modo de desarrollador**
3. Reiniciar Windows

### 3. Configurar el Proyecto

```powershell
# Clonar repositorio
git clone https://github.com/yamil-simon-tsoft/test-app-vb.net.git
cd test-app-vb.net

# Crear entorno virtual
python -m venv venv

# Activar entorno
.\venv\Scripts\Activate.ps1

# Instalar dependencias
pip install -r requirements.txt

# Configurar variables
copy .env.example .env
notepad .env  # Editar con rutas correctas
```

### 4. Configurar .env

Editar `.env` con tus rutas:

```
WINAPPDRIVER_URL=http://127.0.0.1:4723
WORD_APP_PATH=C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE
SCREENSHOT_PATH=reports/screenshots
```

### 5. Verificar Instalación

```powershell
# Iniciar WinAppDriver como Administrador
WinAppDriver.exe

# En otra terminal, ejecutar ejemplo
python examples\word_examples\01_word_basic_operations.py
```

## Solución de Problemas

### WinAppDriver no inicia
- Ejecutar como Administrador
- Verificar que el puerto 4723 esté libre

### Word no encontrado
```powershell
# Buscar ruta de Word
Get-Command winword | Select-Object Source
```

### Permisos denegados
```powershell
# Cambiar política de ejecución
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Próximos Pasos

Una vez instalado, consultar:
- [README.md](../README.md) - Uso básico
- [word_automation_guide.md](word_automation_guide.md) - Guía técnica
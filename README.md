# Automatización de Microsoft Word con WinAppDriver

Framework simple para automatizar Microsoft Word usando Python y WinAppDriver.

## 🚀 Características

- Automatización de Word con WinAppDriver
- Ejemplo funcional de operaciones básicas
- Configuración simple con archivos .env
- Logging y capturas de pantalla automáticas

## 📋 Requisitos

- Windows 10/11
- Python 3.8+
- Microsoft Word 2016+
- WinAppDriver
- Modo de Desarrollador habilitado en Windows

## ⚙️ Instalación

### 1. Instalar WinAppDriver

Descargar desde: https://github.com/Microsoft/WinAppDriver/releases

### 2. Configurar el proyecto

```powershell
# Clonar repositorio
git clone https://github.com/yamil-simon-tsoft/test-app-vb.net.git
cd test-app-vb.net

# Crear entorno virtual
python -m venv venv
.\venv\Scripts\Activate.ps1

# Instalar dependencias
pip install -r requirements.txt

# Configurar variables de entorno
copy .env.example .env
# Editar .env con la ruta correcta de Word
```

## 🎮 Uso

### Ejecutar el ejemplo

```powershell
# 1. Iniciar WinAppDriver como Administrador
WinAppDriver.exe

# 2. Ejecutar el ejemplo
python examples\word_examples\01_word_basic_operations.py
```

## 📁 Estructura del Proyecto

```
proyecto/
├── src/
│   ├── drivers/
│   │   └── winapp_driver.py      # Driver de WinAppDriver
│   └── utils/
│       └── config.py              # Configuración
├── examples/
│   └── word_examples/
│       └── 01_word_basic_operations.py  # Ejemplo de Word
├── reports/
│   ├── logs/                      # Logs de ejecución
│   └── screenshots/               # Capturas de pantalla
├── .env.example                   # Plantilla de configuración
├── requirements.txt               # Dependencias
└── README.md
```

## 📊 Ejemplo de Código

```python
from drivers.winapp_driver import WinAppDriver
from utils.config import config

# Inicializar driver
driver = WinAppDriver(app_path=config.get_word_app_path())

# Iniciar Word
driver.start_driver()
time.sleep(3)

# Navegar ribbon
driver.send_key_combination("alt", "h")  # Tab Inicio

# Tomar captura
driver.take_screenshot("ejemplo")

# Cerrar
driver.send_key_combination("alt", "f4")
driver.stop_driver()
```

## 🔧 Configuración (.env)

```
WINAPPDRIVER_URL=http://127.0.0.1:4723
WORD_APP_PATH=C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE
SCREENSHOT_PATH=reports/screenshots
```

## 📖 Documentación

- [Guía de Instalación](docs/installation_guide.md)
- [Guía de Automatización de Word](docs/word_automation_guide.md)

## 🔍 Solución de Problemas

### Word no encontrado
Verificar la ruta en `.env` y asegurarse que Word esté instalado.

### WinAppDriver no conecta
- Ejecutar WinAppDriver como Administrador
- Verificar que esté corriendo en el puerto 4723

### Modo de Desarrollador
Activar en: Configuración > Actualización y seguridad > Para desarrolladores

## 📞 Soporte

Revisar logs en `reports/logs/` y capturas en `reports/screenshots/`

---

**Desarrollado por**: QA Automation Team - TSOFT  
**Proyecto**: TERNIUM - Automatización VB.NET  
**Versión**: 2.0 (Simplificada)

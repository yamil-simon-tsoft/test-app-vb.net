"""
Configuración centralizada para el proyecto de automatización.

Este módulo maneja todas las configuraciones del proyecto, incluyendo
variables de entorno, rutas de aplicaciones y configuraciones de comportamiento.
"""

import os
import sys
from pathlib import Path
from typing import Optional, Dict, Any
from dotenv import load_dotenv
import logging


class Config:
    """
    Clase de configuración centralizada para automatización de aplicaciones VB.NET.
    """
    
    def __init__(self):
        """Inicializa la configuración cargando variables de entorno."""
        self.project_root = Path(__file__).parent.parent.parent
        self._load_environment_variables()
        self._setup_logging()
        
    def _load_environment_variables(self) -> None:
        """Carga variables de entorno desde archivo .env si existe."""
        env_file = self.project_root / '.env'
        if env_file.exists():
            load_dotenv(env_file)
        
        # WinAppDriver Configuration
        self.WINAPPDRIVER_URL = os.getenv('WINAPPDRIVER_URL', 'http://127.0.0.1:4723')
        self.WINAPPDRIVER_TIMEOUT = int(os.getenv('WINAPPDRIVER_TIMEOUT', '30'))
        
        # Microsoft Word Configuration
        self.WORD_APP_PATH = os.getenv('WORD_APP_PATH', 
            r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE')
        self.WORD_APP_ID = os.getenv('WORD_APP_ID', 'Microsoft.Office.WINWORD')
        
        # VB.NET Application Configuration
        self.VB_APP_PATH = os.getenv('VB_APP_PATH', r'C:\Path\To\Your\VBApp.exe')
        self.VB_APP_ID = os.getenv('VB_APP_ID', 'YourVBApp.Application')
        self.VB_APP_WINDOW_TITLE = os.getenv('VB_APP_WINDOW_TITLE', 'Your VB Application Title')
        
        # Timeouts Configuration
        self.IMPLICIT_WAIT = int(os.getenv('IMPLICIT_WAIT', '10'))
        self.EXPLICIT_WAIT = int(os.getenv('EXPLICIT_WAIT', '20'))
        self.APPLICATION_STARTUP_TIMEOUT = int(os.getenv('APPLICATION_STARTUP_TIMEOUT', '30'))
        
        # Logging Configuration
        self.LOG_LEVEL = os.getenv('LOG_LEVEL', 'INFO').upper()
        self.LOG_FILE_PATH = os.getenv('LOG_FILE_PATH', 'reports/logs/automation.log')
        self.CONSOLE_LOG_ENABLED = os.getenv('CONSOLE_LOG_ENABLED', 'true').lower() == 'true'
        
        # Screenshots Configuration
        self.SCREENSHOT_ON_ERROR = os.getenv('SCREENSHOT_ON_ERROR', 'true').lower() == 'true'
        self.SCREENSHOT_ON_SUCCESS = os.getenv('SCREENSHOT_ON_SUCCESS', 'false').lower() == 'true'
        self.SCREENSHOT_PATH = os.getenv('SCREENSHOT_PATH', 'reports/screenshots')
        self.SCREENSHOT_FORMAT = os.getenv('SCREENSHOT_FORMAT', 'PNG').upper()
        
        # Automation Behavior
        self.AUTO_RETRY_COUNT = int(os.getenv('AUTO_RETRY_COUNT', '3'))
        self.RETRY_DELAY_SECONDS = float(os.getenv('RETRY_DELAY_SECONDS', '2'))
        self.CLOSE_APP_ON_FINISH = os.getenv('CLOSE_APP_ON_FINISH', 'true').lower() == 'true'
        self.PAUSE_BETWEEN_ACTIONS = float(os.getenv('PAUSE_BETWEEN_ACTIONS', '0.5'))
        
        # Word Specific Settings
        self.WORD_STARTUP_TIMEOUT = int(os.getenv('WORD_STARTUP_TIMEOUT', '15'))
        self.WORD_DOCUMENT_SAVE_PATH = os.getenv('WORD_DOCUMENT_SAVE_PATH', 'reports/documents')
        self.WORD_AUTO_SAVE = os.getenv('WORD_AUTO_SAVE', 'true').lower() == 'true'
        
        # Development Settings
        self.DEBUG_MODE = os.getenv('DEBUG_MODE', 'false').lower() == 'true'
        self.VERBOSE_LOGGING = os.getenv('VERBOSE_LOGGING', 'false').lower() == 'true'
        
    def _setup_logging(self) -> None:
        """Configura el sistema de logging."""
        # Convertir string de nivel a constante de logging
        log_level = getattr(logging, self.LOG_LEVEL, logging.INFO)
        
        # Crear directorio de logs si no existe
        log_dir = Path(self.LOG_FILE_PATH).parent
        log_dir.mkdir(parents=True, exist_ok=True)
        
        # Configurar formato de logging
        log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        
        # Configurar logging
        logging.basicConfig(
            level=log_level,
            format=log_format,
            handlers=[
                logging.FileHandler(self.LOG_FILE_PATH, encoding='utf-8'),
                logging.StreamHandler(sys.stdout) if self.CONSOLE_LOG_ENABLED else logging.NullHandler()
            ]
        )
    
    def create_directories(self) -> None:
        """Crea directorios necesarios si no existen."""
        directories = [
            self.SCREENSHOT_PATH,
            Path(self.LOG_FILE_PATH).parent,
            self.WORD_DOCUMENT_SAVE_PATH,
            'reports',
            'temp'
        ]
        
        for directory in directories:
            Path(directory).mkdir(parents=True, exist_ok=True)
    
    # Métodos getter para acceso controlado a la configuración
    
    def get_winappdriver_url(self) -> str:
        """Obtiene la URL de WinAppDriver."""
        return self.WINAPPDRIVER_URL
    
    def get_winappdriver_timeout(self) -> int:
        """Obtiene el timeout de WinAppDriver."""
        return self.WINAPPDRIVER_TIMEOUT
    
    def get_word_app_path(self) -> str:
        """Obtiene la ruta de Microsoft Word."""
        return self.WORD_APP_PATH
    
    def get_word_app_id(self) -> str:
        """Obtiene el ID de aplicación de Word."""
        return self.WORD_APP_ID
    
    def get_vb_app_path(self) -> str:
        """Obtiene la ruta de la aplicación VB.NET."""
        return self.VB_APP_PATH
    
    def get_vb_app_id(self) -> str:
        """Obtiene el ID de la aplicación VB.NET."""
        return self.VB_APP_ID
    
    def get_vb_app_window_title(self) -> str:
        """Obtiene el título de ventana de la aplicación VB.NET."""
        return self.VB_APP_WINDOW_TITLE
    
    def get_implicit_wait(self) -> int:
        """Obtiene el tiempo de espera implícito."""
        return self.IMPLICIT_WAIT
    
    def get_explicit_wait(self) -> int:
        """Obtiene el tiempo de espera explícito."""
        return self.EXPLICIT_WAIT
    
    def get_application_startup_timeout(self) -> int:
        """Obtiene el timeout de inicio de aplicación."""
        return self.APPLICATION_STARTUP_TIMEOUT
    
    def get_screenshot_path(self) -> str:
        """Obtiene la ruta de capturas de pantalla."""
        return self.SCREENSHOT_PATH
    
    def should_screenshot_on_error(self) -> bool:
        """Indica si tomar captura en errores."""
        return self.SCREENSHOT_ON_ERROR
    
    def should_screenshot_on_success(self) -> bool:
        """Indica si tomar captura en éxitos."""
        return self.SCREENSHOT_ON_SUCCESS
    
    def get_retry_count(self) -> int:
        """Obtiene el número de reintentos."""
        return self.AUTO_RETRY_COUNT
    
    def get_retry_delay(self) -> float:
        """Obtiene el delay entre reintentos."""
        return self.RETRY_DELAY_SECONDS
    
    def should_close_app_on_finish(self) -> bool:
        """Indica si cerrar aplicación al finalizar."""
        return self.CLOSE_APP_ON_FINISH
    
    def get_pause_between_actions(self) -> float:
        """Obtiene pausa entre acciones."""
        return self.PAUSE_BETWEEN_ACTIONS
    
    def get_word_startup_timeout(self) -> int:
        """Obtiene timeout de inicio de Word."""
        return self.WORD_STARTUP_TIMEOUT
    
    def get_word_document_save_path(self) -> str:
        """Obtiene ruta de guardado de documentos Word."""
        return self.WORD_DOCUMENT_SAVE_PATH
    
    def is_word_auto_save_enabled(self) -> bool:
        """Indica si auto-guardado de Word está habilitado."""
        return self.WORD_AUTO_SAVE
    
    def is_debug_mode(self) -> bool:
        """Indica si está en modo debug."""
        return self.DEBUG_MODE
    
    def is_verbose_logging(self) -> bool:
        """Indica si logging verboso está habilitado."""
        return self.VERBOSE_LOGGING
    
    def get_word_capabilities(self) -> Dict[str, Any]:
        """
        Obtiene capacidades específicas para Microsoft Word.
        
        Returns:
            Dict con capacidades configuradas para Word
        """
        return {
            'app': self.get_word_app_path(),
            'platformName': 'Windows',
            'deviceName': 'WindowsPC',
            'ms:waitForAppLaunch': self.get_word_startup_timeout(),
            'ms:experimental-webdriver': True,
            'newCommandTimeout': self.get_winappdriver_timeout()
        }
    
    def get_vb_app_capabilities(self) -> Dict[str, Any]:
        """
        Obtiene capacidades específicas para aplicaciones VB.NET.
        
        Returns:
            Dict con capacidades configuradas para VB.NET
        """
        return {
            'app': self.get_vb_app_path(),
            'platformName': 'Windows',
            'deviceName': 'WindowsPC',
            'ms:waitForAppLaunch': self.get_application_startup_timeout(),
            'ms:experimental-webdriver': True,
            'newCommandTimeout': self.get_winappdriver_timeout()
        }
    
    def get_desktop_capabilities(self) -> Dict[str, Any]:
        """
        Obtiene capacidades para conectar al escritorio de Windows.
        
        Returns:
            Dict con capacidades para desktop
        """
        return {
            'app': 'Root',
            'platformName': 'Windows',
            'deviceName': 'WindowsPC',
            'ms:experimental-webdriver': True,
            'newCommandTimeout': self.get_winappdriver_timeout()
        }
    
    def validate_configuration(self) -> Dict[str, bool]:
        """
        Valida la configuración actual.
        
        Returns:
            Dict con resultados de validación
        """
        validation_results = {}
        
        # Validar rutas de aplicaciones
        validation_results['word_path_exists'] = Path(self.get_word_app_path()).exists()
        validation_results['vb_app_path_exists'] = Path(self.get_vb_app_path()).exists()
        
        # Validar directorios
        validation_results['screenshot_dir_writable'] = self._is_directory_writable(self.get_screenshot_path())
        validation_results['log_dir_writable'] = self._is_directory_writable(Path(self.LOG_FILE_PATH).parent)
        
        # Validar configuración de logging
        validation_results['valid_log_level'] = hasattr(logging, self.LOG_LEVEL)
        
        return validation_results
    
    def _is_directory_writable(self, path: Path) -> bool:
        """Verifica si un directorio es escribible."""
        try:
            path = Path(path)
            path.mkdir(parents=True, exist_ok=True)
            test_file = path / 'test_write.tmp'
            test_file.write_text('test')
            test_file.unlink()
            return True
        except Exception:
            return False
    
    def __str__(self) -> str:
        """Representación en string de la configuración."""
        return f"""Configuración de automatización:
- WinAppDriver URL: {self.WINAPPDRIVER_URL}
- Word Path: {self.WORD_APP_PATH}
- VB App Path: {self.VB_APP_PATH}
- Debug Mode: {self.DEBUG_MODE}
- Screenshot on Error: {self.SCREENSHOT_ON_ERROR}
- Log Level: {self.LOG_LEVEL}"""


# Instancia global de configuración
config = Config()
"""
Configuración simple para automatización de Word.
"""
import os
from pathlib import Path
from dotenv import load_dotenv


class Config:
    """Configuración centralizada."""
    
    def __init__(self):
        """Inicializa configuración."""
        self.project_root = Path(__file__).parent.parent.parent
        env_file = self.project_root / '.env'
        if env_file.exists():
            load_dotenv(env_file)
        
        # Configuración básica
        self.WINAPPDRIVER_URL = os.getenv('WINAPPDRIVER_URL', 'http://127.0.0.1:4723')
        self.WORD_APP_PATH = os.getenv('WORD_APP_PATH', 
            r'C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE')
        self.SCREENSHOT_PATH = os.getenv('SCREENSHOT_PATH', 'reports/screenshots')
        
    def create_directories(self):
        """Crea directorios necesarios."""
        directories = [self.SCREENSHOT_PATH, 'reports/logs']
        for directory in directories:
            Path(directory).mkdir(parents=True, exist_ok=True)
    
    def get_winappdriver_url(self):
        """URL de WinAppDriver."""
        return self.WINAPPDRIVER_URL
    
    def get_word_app_path(self):
        """Ruta de Word."""
        return self.WORD_APP_PATH
    
    def get_screenshot_path(self):
        """Ruta de capturas."""
        return self.SCREENSHOT_PATH


# Instancia global
config = Config()
"""
Driver simple para automatización de Microsoft Word usando pywinauto.
"""
import time
import logging
from pathlib import Path

from pywinauto.application import Application
from pywinauto import keyboard

from utils.config import config


class WinAppDriver:
    """Driver simple para automatización de Windows."""
    
    def __init__(self, app_path: str):
        """Inicializa el driver con la ruta de la aplicación."""
        self.app_path = app_path
        self.app = None
        self.logger = logging.getLogger(__name__)
        config.create_directories()
        
    def start_driver(self):
        """Inicia la aplicación."""
        try:
            self.app = Application(backend="uia").start(self.app_path)
            time.sleep(3)  # Esperar que la app se inicie
            self.logger.info("Aplicación iniciada")
            return self.app
        except Exception as e:
            self.logger.error(f"Error: {e}")
            raise
    
    def stop_driver(self):
        """Cierra la aplicación."""
        if self.app:
            try:
                self.app.kill()
            except:
                pass
            self.app = None
    
    def find_elements_by_class_name(self, class_name: str):
        """Encuentra elementos por nombre de clase."""
        try:
            if not self.app:
                return []
            window = self.app.top_window()
            elements = window.descendants(class_name=class_name)
            return elements
        except:
            return []
    
    def take_screenshot(self, filename: str):
        """Toma una captura de pantalla."""
        try:
            screenshot_path = Path(config.get_screenshot_path())
            screenshot_path.mkdir(parents=True, exist_ok=True)
            filepath = screenshot_path / f"{filename}.png"
            
            if self.app:
                window = self.app.top_window()
                window.capture_as_image().save(str(filepath))
                self.logger.info(f"Captura: {filepath}")
        except Exception as e:
            self.logger.warning(f"No se pudo capturar: {e}")
    
    def send_key_combination(self, *keys):
        """Envía combinación de teclas."""
        try:
            # Convertir a formato pywinauto
            key_combo = "+".join(keys)
            keyboard.send_keys(key_combo)
            time.sleep(0.5)
            return True
        except:
            return False
    
    def get_current_window_title(self):
        """Obtiene el título de la ventana."""
        try:
            if self.app:
                return self.app.top_window().window_text()
            return ""
        except:
            return ""

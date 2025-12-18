"""
WinAppDriver simple para automatización de Microsoft Word.
"""
import time
import logging
from pathlib import Path

from appium import webdriver
from appium.options.windows import WindowsOptions
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import WebDriverException

from utils.config import config


class WinAppDriver:
    """Driver simple para WinAppDriver."""
    
    def __init__(self, app_path: str):
        """Inicializa el driver con la ruta de la aplicación."""
        self.app_path = app_path
        self.driver = None
        self.logger = logging.getLogger(__name__)
        config.create_directories()
        
    def start_driver(self):
        """Inicia WinAppDriver."""
        try:
            options = WindowsOptions()
            options.app = self.app_path
            options.platform_name = "Windows"
            options.device_name = "WindowsPC"
            
            self.driver = webdriver.Remote(
                command_executor=config.get_winappdriver_url(),
                options=options
            )
            self.logger.info("WinAppDriver iniciado")
            return self.driver
        except WebDriverException as e:
            self.logger.error(f"Error: {e}")
            raise
    
    def stop_driver(self):
        """Detiene el driver."""
        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None
    
    def find_elements_by_class_name(self, class_name: str):
        """Encuentra elementos por nombre de clase."""
        try:
            time.sleep(1)
            return self.driver.find_elements("class name", class_name)
        except:
            return []
    
    def take_screenshot(self, filename: str):
        """Toma una captura de pantalla."""
        try:
            screenshot_path = Path(config.get_screenshot_path())
            screenshot_path.mkdir(parents=True, exist_ok=True)
            filepath = screenshot_path / f"{filename}.png"
            self.driver.save_screenshot(str(filepath))
            self.logger.info(f"Captura: {filepath}")
        except:
            pass
    
    def send_key_combination(self, *keys):
        """Envía combinación de teclas."""
        try:
            key_combination = ""
            for key in keys:
                if key.lower() == "ctrl":
                    key_combination += Keys.CONTROL
                elif key.lower() == "alt":
                    key_combination += Keys.ALT
                elif key.lower() == "shift":
                    key_combination += Keys.SHIFT
                elif key.lower() == "enter":
                    key_combination += Keys.ENTER
                else:
                    key_combination += key
            
            active_element = self.driver.switch_to.active_element
            if active_element:
                active_element.send_keys(key_combination)
            time.sleep(0.5)
            return True
        except:
            return False
    
    def get_current_window_title(self):
        """Obtiene el título de la ventana."""
        try:
            return self.driver.title or ""
        except:
            return ""

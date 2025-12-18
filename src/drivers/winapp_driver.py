"""
WinAppDriver para automatización de aplicaciones Windows.

Este módulo proporciona un wrapper completo para WinAppDriver que simplifica
la automatización de aplicaciones de escritorio Windows, especialmente
diseñado para aplicaciones VB.NET y Microsoft Office.
"""

import os
import sys
import time
import logging
from typing import Optional, Dict, Any, Tuple, List
from pathlib import Path

try:
    from appium import webdriver
    from appium.options.windows import WindowsOptions
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.common.exceptions import (
        WebDriverException,
        TimeoutException,
        NoSuchElementException,
        ElementNotInteractableException,
        InvalidSessionIdException
    )
except ImportError as e:
    print(f"Error al importar dependencias requeridas: {e}")
    print("Por favor ejecute: pip install appium-python-client selenium")
    sys.exit(1)

# Importar configuración del proyecto
from utils.config import config


class WinAppDriverError(Exception):
    """Excepción personalizada para errores de WinAppDriver."""
    pass


class WinAppDriver:
    """
    Clase principal para manejar WinAppDriver y automatización de aplicaciones Windows.
    """
    
    def __init__(self, app_path: Optional[str] = None, app_id: Optional[str] = None):
        """
        Inicializa el driver de WinAppDriver.
        
        Args:
            app_path: Ruta opcional a la aplicación a automatizar
            app_id: ID opcional de la aplicación (para apps ya ejecutándose)
        """
        self.driver: Optional[webdriver.Remote] = None
        self.wait: Optional[WebDriverWait] = None
        self.app_path = app_path
        self.app_id = app_id
        self.logger = logging.getLogger(__name__)
        self.is_connected = False
        
        # Configurar directorios necesarios
        config.create_directories()
        
    def start_driver(self, capabilities: Optional[Dict[str, Any]] = None) -> webdriver.Remote:
        """
        Inicia el driver WinAppDriver con capacidades específicas.
        
        Args:
            capabilities: Capacidades personalizadas para el driver
            
        Returns:
            Instancia del driver WebDriver
            
        Raises:
            WinAppDriverError: Si no se puede iniciar el driver
        """
        try:
            self.logger.info("Iniciando WinAppDriver...")
            
            # Configurar opciones del driver
            options = WindowsOptions()
            
            # Configurar aplicación objetivo
            if self.app_path:
                options.app = self.app_path
                self.logger.info(f"Configurado para aplicación: {self.app_path}")
            elif self.app_id:
                options.app = self.app_id
                self.logger.info(f"Configurado para ID de aplicación: {self.app_id}")
            else:
                # Conectar al escritorio de Windows
                options.app = "Root"
                self.logger.info("Configurado para escritorio de Windows")
            
            # Aplicar capacidades adicionales
            if capabilities:
                for key, value in capabilities.items():
                    setattr(options, key, value)
                    self.logger.debug(f"Capacidad configurada: {key} = {value}")
            
            # Configurar capacidades por defecto
            options.platform_name = "Windows"
            options.device_name = "WindowsPC"
            
            # Inicializar el driver
            self.driver = webdriver.Remote(
                command_executor=config.get_winappdriver_url(),
                options=options
            )
            
            # Configurar timeouts
            self.driver.implicitly_wait(config.get_implicit_wait())
            self.wait = WebDriverWait(self.driver, config.get_explicit_wait())
            
            self.is_connected = True
            self.logger.info("WinAppDriver iniciado exitosamente")
            
            # Tomar captura inicial si está habilitado
            if config.should_screenshot_on_success():
                self.take_screenshot("driver_started")
            
            return self.driver
            
        except WebDriverException as e:
            error_msg = f"Error al conectar con WinAppDriver: {str(e)}"
            self.logger.error(error_msg)
            self.logger.error("Verifique que WinAppDriver esté ejecutándose como administrador")
            raise WinAppDriverError(error_msg) from e
        
        except Exception as e:
            error_msg = f"Error inesperado al iniciar WinAppDriver: {str(e)}"
            self.logger.error(error_msg)
            raise WinAppDriverError(error_msg) from e
    
    def stop_driver(self) -> None:
        """Detiene el driver WinAppDriver de forma segura."""
        try:
            if self.driver and self.is_connected:
                self.logger.info("Deteniendo WinAppDriver...")
                self.driver.quit()
                self.logger.info("WinAppDriver detenido exitosamente")
        except Exception as e:
            self.logger.warning(f"Error al detener WinAppDriver: {str(e)}")
        finally:
            self.driver = None
            self.wait = None
            self.is_connected = False
    
    def find_element_with_retry(self, locator: Tuple[str, str], 
                               timeout: int = None, 
                               retry_count: int = None) -> Any:
        """
        Encuentra un elemento con reintentos automáticos.
        
        Args:
            locator: Tupla (tipo_localizador, valor)
            timeout: Tiempo de espera personalizado
            retry_count: Número de reintentos personalizados
            
        Returns:
            WebElement encontrado
            
        Raises:
            NoSuchElementException: Si no se encuentra el elemento
        """
        timeout = timeout or config.get_explicit_wait()
        retry_count = retry_count or config.get_retry_count()
        
        for attempt in range(retry_count + 1):
            try:
                self.logger.debug(f"Buscando elemento {locator} - Intento {attempt + 1}")
                
                wait = WebDriverWait(self.driver, timeout)
                element = wait.until(EC.presence_of_element_located(locator))
                
                self.logger.debug(f"Elemento encontrado: {locator}")
                return element
                
            except TimeoutException:
                if attempt < retry_count:
                    self.logger.warning(f"Elemento no encontrado en intento {attempt + 1}, reintentando...")
                    time.sleep(config.get_retry_delay())
                else:
                    error_msg = f"Elemento no encontrado después de {retry_count + 1} intentos: {locator}"
                    self.logger.error(error_msg)
                    
                    if config.should_screenshot_on_error():
                        self.take_screenshot(f"element_not_found_{int(time.time())}")
                    
                    raise NoSuchElementException(error_msg)
    
    def find_element_by_automation_id(self, automation_id: str, timeout: int = None) -> Any:
        """Encuentra elemento por AutomationId."""
        return self.find_element_with_retry((By.NAME, automation_id), timeout)
    
    def find_element_by_name(self, name: str, timeout: int = None) -> Any:
        """Encuentra elemento por nombre."""
        return self.find_element_with_retry((By.NAME, name), timeout)
    
    def find_element_by_class_name(self, class_name: str, timeout: int = None) -> Any:
        """Encuentra elemento por clase."""
        return self.find_element_with_retry((By.CLASS_NAME, class_name), timeout)
    
    def find_element_by_accessibility_id(self, accessibility_id: str, timeout: int = None) -> Any:
        """Encuentra elemento por AccessibilityId."""
        return self.find_element_with_retry((By.ACCESSIBILITY_ID, accessibility_id), timeout)
    
    def find_elements_by_class_name(self, class_name: str, timeout: int = None) -> List[Any]:
        """Encuentra múltiples elementos por clase."""
        timeout = timeout or config.get_explicit_wait()
        try:
            wait = WebDriverWait(self.driver, timeout)
            elements = wait.until(EC.presence_of_all_elements_located((By.CLASS_NAME, class_name)))
            self.logger.debug(f"Encontrados {len(elements)} elementos con clase: {class_name}")
            return elements
        except TimeoutException:
            self.logger.warning(f"No se encontraron elementos con clase: {class_name}")
            return []
    
    def click_element(self, element_or_locator, timeout: int = None) -> bool:
        """
        Hace clic en un elemento con manejo de errores.
        
        Args:
            element_or_locator: WebElement o tupla de localizador
            timeout: Timeout personalizado
            
        Returns:
            True si el clic fue exitoso
        """
        try:
            if isinstance(element_or_locator, tuple):
                element = self.find_element_with_retry(element_or_locator, timeout)
            else:
                element = element_or_locator
            
            # Esperar a que el elemento sea clickeable
            wait = WebDriverWait(self.driver, timeout or config.get_explicit_wait())
            clickable_element = wait.until(EC.element_to_be_clickable(element))
            
            clickable_element.click()
            self.logger.debug(f"Clic exitoso en elemento")
            
            # Pausa entre acciones si está configurado
            if config.get_pause_between_actions() > 0:
                time.sleep(config.get_pause_between_actions())
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error al hacer clic en elemento: {str(e)}")
            
            if config.should_screenshot_on_error():
                self.take_screenshot(f"click_error_{int(time.time())}")
            
            return False
    
    def send_keys_to_element(self, element_or_locator, text: str, 
                           clear_first: bool = False, timeout: int = None) -> bool:
        """
        Envía texto a un elemento.
        
        Args:
            element_or_locator: WebElement o tupla de localizador
            text: Texto a enviar
            clear_first: Si limpiar el campo primero
            timeout: Timeout personalizado
            
        Returns:
            True si el envío fue exitoso
        """
        try:
            if isinstance(element_or_locator, tuple):
                element = self.find_element_with_retry(element_or_locator, timeout)
            else:
                element = element_or_locator
            
            if clear_first:
                element.clear()
            
            element.send_keys(text)
            self.logger.debug(f"Texto enviado exitosamente: {text[:50]}...")
            
            # Pausa entre acciones
            if config.get_pause_between_actions() > 0:
                time.sleep(config.get_pause_between_actions())
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error al enviar texto: {str(e)}")
            
            if config.should_screenshot_on_error():
                self.take_screenshot(f"send_keys_error_{int(time.time())}")
            
            return False
    
    def get_element_text(self, element_or_locator, timeout: int = None) -> str:
        """
        Obtiene el texto de un elemento.
        
        Args:
            element_or_locator: WebElement o tupla de localizador
            timeout: Timeout personalizado
            
        Returns:
            Texto del elemento
        """
        try:
            if isinstance(element_or_locator, tuple):
                element = self.find_element_with_retry(element_or_locator, timeout)
            else:
                element = element_or_locator
            
            text = element.text or element.get_attribute("Name") or ""
            self.logger.debug(f"Texto obtenido: {text[:50]}...")
            return text
            
        except Exception as e:
            self.logger.error(f"Error al obtener texto: {str(e)}")
            return ""
    
    def is_element_visible(self, locator: Tuple[str, str], timeout: int = 5) -> bool:
        """
        Verifica si un elemento es visible.
        
        Args:
            locator: Tupla de localizador
            timeout: Timeout para la verificación
            
        Returns:
            True si el elemento es visible
        """
        try:
            wait = WebDriverWait(self.driver, timeout)
            wait.until(EC.visibility_of_element_located(locator))
            return True
        except TimeoutException:
            return False
    
    def wait_for_window_title(self, title: str, timeout: int = None) -> bool:
        """
        Espera a que aparezca una ventana con el título especificado.
        
        Args:
            title: Título de la ventana (puede ser parcial)
            timeout: Tiempo de espera
            
        Returns:
            True si la ventana apareció
        """
        timeout = timeout or config.get_explicit_wait()
        
        try:
            wait = WebDriverWait(self.driver, timeout)
            wait.until(EC.title_contains(title))
            self.logger.info(f"Ventana encontrada con título que contiene: {title}")
            return True
        except TimeoutException:
            self.logger.warning(f"No se encontró ventana con título: {title}")
            return False
    
    def switch_to_window_by_title(self, title: str) -> bool:
        """
        Cambia a una ventana específica por su título.
        
        Args:
            title: Título de la ventana (puede ser parcial)
            
        Returns:
            True si se cambió exitosamente
        """
        try:
            current_handles = self.driver.window_handles
            
            for handle in current_handles:
                self.driver.switch_to.window(handle)
                if title.lower() in self.driver.title.lower():
                    self.logger.info(f"Cambiado a ventana: {self.driver.title}")
                    return True
            
            self.logger.warning(f"No se encontró ventana con título: {title}")
            return False
            
        except Exception as e:
            self.logger.error(f"Error al cambiar de ventana: {str(e)}")
            return False
    
    def take_screenshot(self, filename: Optional[str] = None) -> str:
        """
        Toma una captura de pantalla.
        
        Args:
            filename: Nombre personalizado del archivo
            
        Returns:
            Ruta del archivo de captura
        """
        try:
            if not filename:
                timestamp = int(time.time())
                filename = f"screenshot_{timestamp}.png"
            
            # Asegurar extensión
            if not filename.endswith('.png'):
                filename += '.png'
            
            screenshot_path = Path(config.get_screenshot_path())
            screenshot_path.mkdir(parents=True, exist_ok=True)
            
            filepath = screenshot_path / filename
            
            success = self.driver.save_screenshot(str(filepath))
            
            if success:
                self.logger.info(f"Captura guardada: {filepath}")
                return str(filepath)
            else:
                self.logger.warning("No se pudo guardar la captura de pantalla")
                return ""
                
        except Exception as e:
            self.logger.error(f"Error al tomar captura: {str(e)}")
            return ""
    
    def send_key_combination(self, *keys) -> bool:
        """
        Envía una combinación de teclas al elemento activo.
        
        Args:
            *keys: Secuencia de teclas a combinar
            
        Returns:
            True si fue exitoso
        """
        try:
            # Construir combinación de teclas
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
                elif key.lower() == "escape":
                    key_combination += Keys.ESCAPE
                else:
                    key_combination += key
            
            # Enviar a elemento activo o al body
            active_element = self.driver.switch_to.active_element
            if active_element:
                active_element.send_keys(key_combination)
            else:
                self.driver.find_element(By.TAG_NAME, "body").send_keys(key_combination)
            
            self.logger.debug(f"Combinación de teclas enviada: {' + '.join(keys)}")
            
            # Pausa entre acciones
            if config.get_pause_between_actions() > 0:
                time.sleep(config.get_pause_between_actions())
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error al enviar combinación de teclas: {str(e)}")
            return False
    
    def get_current_window_title(self) -> str:
        """
        Obtiene el título de la ventana actual.
        
        Returns:
            Título de la ventana actual
        """
        try:
            return self.driver.title or ""
        except Exception as e:
            self.logger.error(f"Error al obtener título de ventana: {str(e)}")
            return ""
    
    def close_application(self) -> None:
        """Cierra la aplicación de forma segura."""
        try:
            if self.driver and self.is_connected:
                self.logger.info("Cerrando aplicación...")
                
                # Intentar cerrar usando Alt+F4
                self.send_key_combination("alt", "f4")
                time.sleep(2)
                
                # Verificar si la aplicación se cerró
                try:
                    # Si podemos obtener el título, la aplicación aún está abierta
                    title = self.driver.title
                    if title:
                        self.logger.warning("Aplicación aún abierta, forzando cierre...")
                        self.driver.close()
                except InvalidSessionIdException:
                    # La aplicación se cerró correctamente
                    self.logger.info("Aplicación cerrada exitosamente")
                
        except Exception as e:
            self.logger.warning(f"Error al cerrar aplicación: {str(e)}")
        
        finally:
            self.stop_driver()
    
    def __enter__(self):
        """Soporte para context manager."""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Limpieza automática al salir del context manager."""
        if config.should_close_app_on_finish():
            self.close_application()
        else:
            self.stop_driver()
        
        # Log de errores si ocurrieron
        if exc_type:
            self.logger.error(f"Error en context manager: {exc_type.__name__}: {exc_val}")
            
            if config.should_screenshot_on_error():
                self.take_screenshot(f"context_error_{int(time.time())}")
    
    def __del__(self):
        """Destructor para limpieza final."""
        if hasattr(self, 'is_connected') and self.is_connected:
            self.stop_driver()
        except Exception as e:
            self.logger.warning(f"Error al cerrar aplicación: {str(e)}")
        
        finally:
            self.stop_driver()
    
    def __enter__(self):
        """Soporte para context manager."""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Limpieza automática al salir del context manager."""
        if config.should_close_app_on_finish():
            self.close_application()
        else:
            self.stop_driver()
        
        # Log de errores si ocurrieron
        if exc_type:
            self.logger.error(f"Error en context manager: {exc_type.__name__}: {exc_val}")
            
            if config.should_screenshot_on_error():
                self.take_screenshot(f"context_error_{int(time.time())}")
    
    def __del__(self):
        """Destructor para limpieza final."""
        if hasattr(self, 'is_connected') and self.is_connected:
            self.stop_driver()
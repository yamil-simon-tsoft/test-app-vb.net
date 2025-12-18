"""
Ejemplo 1: Operaciones Básicas de Microsoft Word

Este ejemplo demuestra cómo:
1. Iniciar Microsoft Word usando WinAppDriver
2. Verificar que la aplicación se abrió correctamente
3. Navegar por la interfaz principal
4. Cerrar la aplicación de forma segura

Autor: QA Automation Team
Fecha: Diciembre 2025
"""

import sys
import time
import logging
from pathlib import Path

# Agregar src al path para importaciones
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root / "src"))

from drivers.winapp_driver import WinAppDriver, WinAppDriverError
from utils.config import config


class WordBasicOperationsExample:
    """
    Clase de ejemplo para operaciones básicas de Microsoft Word.
    """
    
    def __init__(self):
        """Inicializa el ejemplo."""
        self.logger = logging.getLogger(__name__)
        self.driver = None
        
    def setup(self) -> bool:
        """
        Configuración inicial del ejemplo.
        
        Returns:
            True si la configuración fue exitosa
        """
        try:
            self.logger.info("=== INICIANDO EJEMPLO: OPERACIONES BÁSICAS DE WORD ===")
            
            # Crear directorios necesarios
            config.create_directories()
            
            # Inicializar driver con capacidades específicas para Word
            self.driver = WinAppDriver(app_path=config.get_word_app_path())
            
            return True
            
        except Exception as e:
            self.logger.error(f"Error en configuración inicial: {str(e)}")
            return False
    
    def start_word_application(self) -> bool:
        """
        Paso 1: Iniciar Microsoft Word.
        
        Returns:
            True si Word se inició exitosamente
        """
        try:
            self.logger.info("PASO 1: Iniciando Microsoft Word...")
            
            # Configurar capacidades específicas para Word
            word_capabilities = config.get_word_capabilities()
            
            # Iniciar el driver
            self.driver.start_driver(word_capabilities)
            
            # Esperar a que Word se cargue completamente
            self.logger.info("Esperando a que Word se cargue completamente...")
            time.sleep(3)
            
            # Verificar que Word se inició correctamente
            current_title = self.driver.get_current_window_title()
            self.logger.info(f"Título de ventana actual: {current_title}")
            
            if "Word" in current_title or "Microsoft" in current_title:
                self.logger.info("✅ Microsoft Word iniciado exitosamente")
                
                # Tomar captura de pantalla del estado inicial
                screenshot_path = self.driver.take_screenshot("word_started_successfully")
                self.logger.info(f"Captura inicial guardada: {screenshot_path}")
                
                return True
            else:
                self.logger.error("❌ No se detectó Microsoft Word en la ventana actual")
                return False
                
        except WinAppDriverError as e:
            self.logger.error(f"❌ Error específico de WinAppDriver: {str(e)}")
            self.logger.error("Verifique que:")
            self.logger.error("1. WinAppDriver esté ejecutándose como administrador")
            self.logger.error("2. Microsoft Word esté instalado en la ruta configurada")
            self.logger.error("3. Developer Mode esté habilitado en Windows")
            return False
            
        except Exception as e:
            self.logger.error(f"❌ Error inesperado al iniciar Word: {str(e)}")
            return False
    
    def verify_word_interface(self) -> bool:
        """
        Paso 2: Verificar elementos principales de la interfaz de Word.
        
        Returns:
            True si se verificaron los elementos correctamente
        """
        try:
            self.logger.info("PASO 2: Verificando interfaz principal de Word...")
            
            # Lista de elementos a verificar en la interfaz de Word
            elements_to_verify = [
                ("Ribbon", "NetUIRibbonTab"),  # Pestañas del ribbon
                ("Document Area", "_WwG"),      # Área del documento
                ("Title Bar", "TitleBar"),      # Barra de título
            ]
            
            verification_results = {}
            
            for element_name, class_name in elements_to_verify:
                try:
                    self.logger.info(f"Verificando {element_name}...")
                    
                    # Buscar elemento con timeout corto
                    elements = self.driver.find_elements_by_class_name(class_name, timeout=5)
                    
                    if elements:
                        verification_results[element_name] = True
                        self.logger.info(f"✅ {element_name} encontrado ({len(elements)} elementos)")
                    else:
                        verification_results[element_name] = False
                        self.logger.warning(f"⚠️ {element_name} no encontrado")
                        
                except Exception as e:
                    verification_results[element_name] = False
                    self.logger.warning(f"⚠️ Error verificando {element_name}: {str(e)}")
            
            # Evaluar resultados
            successful_verifications = sum(verification_results.values())
            total_verifications = len(verification_results)
            
            self.logger.info(f"Verificaciones exitosas: {successful_verifications}/{total_verifications}")
            
            if successful_verifications >= 2:  # Al menos 2 de 3 elementos encontrados
                self.logger.info("✅ Interfaz de Word verificada correctamente")
                
                # Tomar captura de la interfaz verificada
                self.driver.take_screenshot("word_interface_verified")
                
                return True
            else:
                self.logger.error("❌ No se pudieron verificar suficientes elementos de la interfaz")
                return False
                
        except Exception as e:
            self.logger.error(f"❌ Error al verificar interfaz: {str(e)}")
            return False
    
    def navigate_ribbon_tabs(self) -> bool:
        """
        Paso 3: Navegar por las pestañas principales del ribbon.
        
        Returns:
            True si se navegó exitosamente
        """
        try:
            self.logger.info("PASO 3: Navegando por pestañas del ribbon...")
            
            # Lista de pestañas comunes del ribbon (usando teclas de acceso rápido)
            ribbon_tabs = [
                ("Inicio", "h"),
                ("Insertar", "n"),
                ("Diseño", "g"),
                ("Referencias", "s"),
                ("Correspondencia", "m"),
                ("Revisar", "r"),
                ("Vista", "w")
            ]
            
            navigation_success = 0
            
            for tab_name, access_key in ribbon_tabs:
                try:
                    self.logger.info(f"Navegando a pestaña: {tab_name}")
                    
                    # Usar tecla de acceso rápido Alt + tecla
                    success = self.driver.send_key_combination("alt", access_key)
                    
                    if success:
                        # Esperar un momento para que se cargue la pestaña
                        time.sleep(1)
                        
                        # Tomar captura de la pestaña
                        screenshot_name = f"ribbon_tab_{tab_name.lower()}"
                        self.driver.take_screenshot(screenshot_name)
                        
                        navigation_success += 1
                        self.logger.info(f"✅ Navegación a {tab_name} exitosa")
                    else:
                        self.logger.warning(f"⚠️ No se pudo navegar a {tab_name}")
                    
                    # Pequeña pausa entre navegaciones
                    time.sleep(0.5)
                    
                except Exception as e:
                    self.logger.warning(f"⚠️ Error navegando a {tab_name}: {str(e)}")
            
            # Evaluar éxito de navegación
            if navigation_success >= 3:  # Al menos 3 pestañas navegadas
                self.logger.info(f"✅ Navegación de ribbon exitosa ({navigation_success} pestañas)")
                return True
            else:
                self.logger.warning(f"⚠️ Navegación parcial ({navigation_success} pestañas)")
                return False
                
        except Exception as e:
            self.logger.error(f"❌ Error en navegación de ribbon: {str(e)}")
            return False
    
    def demonstrate_basic_interactions(self) -> bool:
        """
        Paso 4: Demostrar interacciones básicas con Word.
        
        Returns:
            True si las interacciones fueron exitosas
        """
        try:
            self.logger.info("PASO 4: Demostrando interacciones básicas...")
            
            # Ir a la pestaña Inicio
            self.logger.info("Yendo a pestaña Inicio...")
            self.driver.send_key_combination("alt", "h")
            time.sleep(1)
            
            # Intentar escribir texto en el documento
            self.logger.info("Escribiendo texto de ejemplo en el documento...")
            
            # Enviar texto directamente
            sample_text = "Este es un texto de ejemplo para demostrar automatización de Word."
            success = self.driver.send_key_combination("ctrl", "home")  # Ir al inicio del documento
            
            if success:
                time.sleep(0.5)
                
                # Escribir el texto (usando el método más directo)
                from selenium.webdriver.common.keys import Keys
                active_element = self.driver.driver.switch_to.active_element
                if active_element:
                    active_element.send_keys(sample_text)
                    self.logger.info("✅ Texto escrito exitosamente")
                    
                    # Tomar captura con el texto
                    self.driver.take_screenshot("text_written_in_document")
                    
                    # Seleccionar todo el texto
                    time.sleep(1)
                    self.driver.send_key_combination("ctrl", "a")
                    self.logger.info("✅ Texto seleccionado")
                    
                    # Tomar captura con texto seleccionado
                    self.driver.take_screenshot("text_selected")
                    
                    return True
                else:
                    self.logger.warning("⚠️ No se pudo acceder al área de documento")
                    return False
            else:
                self.logger.warning("⚠️ No se pudo posicionar el cursor")
                return False
                
        except Exception as e:
            self.logger.error(f"❌ Error en interacciones básicas: {str(e)}")
            return False
    
    def close_word_safely(self) -> bool:
        """
        Paso 5: Cerrar Microsoft Word de forma segura.
        
        Returns:
            True si se cerró exitosamente
        """
        try:
            self.logger.info("PASO 5: Cerrando Microsoft Word...")
            
            # Intentar cerrar sin guardar (ya que es solo una demostración)
            self.logger.info("Cerrando documento sin guardar...")
            
            # Usar Alt+F4 para cerrar
            success = self.driver.send_key_combination("alt", "f4")
            
            if success:
                # Esperar posible diálogo de guardar
                time.sleep(2)
                
                # Si aparece diálogo de guardar, presionar "N" para No guardar
                try:
                    # Intentar presionar "N" por si aparece el diálogo
                    self.driver.driver.switch_to.active_element.send_keys("n")
                    self.logger.info("Diálogo de guardar detectado - seleccionado No guardar")
                    time.sleep(1)
                except:
                    # No hay diálogo, continuar
                    pass
                
                self.logger.info("✅ Microsoft Word cerrado exitosamente")
                return True
            else:
                self.logger.warning("⚠️ No se pudo enviar comando de cierre")
                return False
                
        except Exception as e:
            self.logger.error(f"❌ Error al cerrar Word: {str(e)}")
            return False
        
        finally:
            # Asegurar que el driver se detenga
            if self.driver:
                self.driver.stop_driver()
    
    def run_complete_example(self) -> bool:
        """
        Ejecuta el ejemplo completo de operaciones básicas de Word.
        
        Returns:
            True si todo el ejemplo se ejecutó exitosamente
        """
        try:
            # Configuración inicial
            if not self.setup():
                return False
            
            steps_results = []
            
            # Ejecutar todos los pasos
            steps_results.append(self.start_word_application())
            
            if steps_results[-1]:  # Solo continuar si el paso anterior fue exitoso
                steps_results.append(self.verify_word_interface())
            
            if steps_results[-1]:
                steps_results.append(self.navigate_ribbon_tabs())
            
            if steps_results[-1]:
                steps_results.append(self.demonstrate_basic_interactions())
            
            # Siempre intentar cerrar Word
            close_result = self.close_word_safely()
            steps_results.append(close_result)
            
            # Evaluar resultados
            successful_steps = sum(steps_results)
            total_steps = len(steps_results)
            
            self.logger.info(f"=== RESUMEN DEL EJEMPLO ===")
            self.logger.info(f"Pasos exitosos: {successful_steps}/{total_steps}")
            
            if successful_steps >= 4:  # Al menos 4 de 5 pasos exitosos
                self.logger.info("🎉 EJEMPLO COMPLETADO EXITOSAMENTE")
                return True
            else:
                self.logger.warning("⚠️ EJEMPLO COMPLETADO CON ADVERTENCIAS")
                return False
                
        except Exception as e:
            self.logger.error(f"❌ ERROR GENERAL EN EJEMPLO: {str(e)}")
            return False


def main():
    """Función principal para ejecutar el ejemplo."""
    
    # Configurar logging detallado para el ejemplo
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler('reports/logs/word_basic_operations.log', encoding='utf-8')
        ]
    )
    
    logger = logging.getLogger(__name__)
    
    try:
        logger.info("🚀 INICIANDO EJEMPLO: OPERACIONES BÁSICAS DE MICROSOFT WORD")
        logger.info("=" * 60)
        
        # Verificar configuración
        logger.info("Verificando configuración...")
        validation_results = config.validate_configuration()
        
        if not validation_results.get('word_path_exists', False):
            logger.error("❌ Microsoft Word no encontrado en la ruta configurada")
            logger.error(f"Ruta configurada: {config.get_word_app_path()}")
            logger.error("Por favor, edite el archivo .env con la ruta correcta de Word")
            return False
        
        logger.info("✅ Configuración válida")
        
        # Crear y ejecutar ejemplo
        example = WordBasicOperationsExample()
        result = example.run_complete_example()
        
        if result:
            logger.info("🎉 EJEMPLO EJECUTADO EXITOSAMENTE")
            logger.info("Revise las capturas de pantalla en: reports/screenshots/")
        else:
            logger.error("❌ EJEMPLO EJECUTADO CON ERRORES")
            logger.error("Revise los logs para más detalles")
        
        return result
        
    except KeyboardInterrupt:
        logger.warning("⚠️ Ejemplo interrumpido por el usuario")
        return False
        
    except Exception as e:
        logger.error(f"❌ Error inesperado en ejemplo: {str(e)}")
        return False
    
    finally:
        logger.info("=" * 60)
        logger.info("FIN DEL EJEMPLO")


if __name__ == "__main__":
    # Ejecutar el ejemplo
    success = main()
    
    # Salir con código apropiado
    sys.exit(0 if success else 1)
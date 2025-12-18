"""
Ejemplo: Operaciones Básicas de Microsoft Word con WinAppDriver

Demuestra:
1. Iniciar Word
2. Verificar interfaz
3. Navegar por ribbon
4. Cerrar Word
"""
import sys
import time
import logging
from pathlib import Path

# Agregar src al path
project_root = Path(__file__).parent.parent.parent
sys.path.insert(0, str(project_root / "src"))

from drivers.winapp_driver import WinAppDriver
from utils.config import config


def main():
    """Ejecuta el ejemplo de Word."""
    
    # Configurar logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler('reports/logs/word_example.log', encoding='utf-8')
        ]
    )
    
    logger = logging.getLogger(__name__)
    logger.info("=== EJEMPLO: WORD BÁSICO ===")
    
    # Crear directorios
    config.create_directories()
    
    # Inicializar driver
    driver = WinAppDriver(app_path=config.get_word_app_path())
    
    try:
        # 1. Iniciar Word
        logger.info("1. Iniciando Word...")
        driver.start_driver()
        time.sleep(3)
        
        title = driver.get_current_window_title()
        logger.info(f"   Título: {title}")
        
        if "Word" not in title:
            logger.error("   Word no detectado")
            return False
        
        driver.take_screenshot("word_started")
        logger.info("   ✓ Word iniciado")
        
        # 2. Verificar interfaz
        logger.info("2. Verificando interfaz...")
        elements = driver.find_elements_by_class_name("NetUIRibbonTab")
        logger.info(f"   Encontrados {len(elements)} tabs del ribbon")
        driver.take_screenshot("word_interface")
        logger.info("   ✓ Interfaz verificada")
        
        # 3. Navegar por ribbon
        logger.info("3. Navegando ribbon...")
        tabs = [("h", "Inicio"), ("n", "Insertar"), ("g", "Diseño")]
        
        for key, name in tabs:
            driver.send_key_combination("alt", key)
            time.sleep(1)
            driver.take_screenshot(f"ribbon_{name.lower()}")
            logger.info(f"   ✓ Tab {name}")
        
        # 4. Cerrar Word
        logger.info("4. Cerrando Word...")
        driver.send_key_combination("alt", "f4")
        time.sleep(2)
        
        # Manejar diálogo "No guardar"
        try:
            driver.driver.switch_to.active_element.send_keys("n")
        except:
            pass
        
        logger.info("   ✓ Word cerrado")
        logger.info("=== EJEMPLO COMPLETADO ===")
        return True
        
    except Exception as e:
        logger.error(f"Error: {e}")
        return False
        
    finally:
        driver.stop_driver()


if __name__ == "__main__":
    success = main()
    sys.exit(0 if success else 1)

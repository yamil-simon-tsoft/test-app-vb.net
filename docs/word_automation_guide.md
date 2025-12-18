# Guía de Automatización de Microsoft Word

## Introducción

Esta guía proporciona información detallada sobre cómo automatizar Microsoft Word usando WinAppDriver y Python. Se basa en el ejemplo implementado `01_word_basic_operations.py` y explica los patrones, mejores prácticas y técnicas utilizadas.

## Conceptos Fundamentales

### Arquitectura de la Automatización

```
Usuario → Python Script → WinAppDriver → Windows UI Automation → Word Application
```

1. **Python Script**: Lógica de automatización
2. **WinAppDriver**: Servidor que expone API de automatización
3. **Windows UI Automation**: Framework nativo de Windows
4. **Word Application**: Aplicación objetivo

### Elementos de Word y Identificación

#### Métodos de Localización
- **Por nombre**: `find_element_by_name("Guardar")`
- **Por ID de automatización**: `find_element_by_accessibility_id("SaveButton")`
- **Por clase**: `find_element_by_class_name("NetUIRibbonButton")`
- **Por XPath**: `find_element_by_xpath("//Button[@Name='Negrita']")`

#### Jerarquía de Elementos
```
Word Application
├── Ribbon (Cinta)
│   ├── Tabs (Pestañas)
│   ├── Groups (Grupos)
│   └── Buttons (Botones)
├── Document Area (Área de documento)
└── Status Bar (Barra de estado)
```

## Patrones de Automatización

### 1. Patrón de Inicialización Segura

```python
class WordAutomation:
    def __init__(self):
        self.driver = None
        self.logger = logging.getLogger(__name__)
    
    def setup(self):
        try:
            self.driver = WinAppDriver(app_path=WORD_PATH)
            self.driver.start_driver(capabilities)
            time.sleep(3)  # Espera estabilización
            
            if not self.verify_word_started():
                raise Exception("Word no inició correctamente")
                
        except Exception as e:
            self.cleanup()
            raise e
    
    def cleanup(self):
        if self.driver:
            self.driver.stop_driver()
```

### 2. Patrón de Verificación de Estado

```python
def verify_word_started(self) -> bool:
    try:
        title = self.driver.get_current_window_title()
        return "Word" in title or "Microsoft Word" in title
    except:
        return False

def wait_for_document_ready(self):
    for _ in range(10):
        try:
            # Verificar que el área de documento esté disponible
            active_element = self.driver.driver.switch_to.active_element
            if active_element:
                return True
        except:
            pass
        time.sleep(1)
    return False
```

### 3. Patrón de Manejo de Errores con Reintentos

```python
def execute_with_retry(self, action, max_retries=3):
    for attempt in range(max_retries):
        try:
            return action()
        except Exception as e:
            if attempt == max_retries - 1:
                self.logger.error(f"Fallo después de {max_retries} intentos: {e}")
                raise e
            
            self.logger.warning(f"Intento {attempt + 1} falló: {e}")
            time.sleep(2)
    return None
```

## Análisis del Ejemplo Implementado

El archivo `01_word_basic_operations.py` implementa estos patrones:

### Patrón de Inicialización Usado
```python
def setup_and_start(self) -> bool:
    config.create_directories()
    self.driver = WinAppDriver(app_path=config.get_word_app_path())
    self.driver.start_driver(config.get_word_capabilities())
    time.sleep(3)  # Estabilización crítica
    return "Word" in self.driver.get_current_window_title()
```

### Patrón de Navegación Implementado
```python
def navigate_ribbon_tabs(self) -> bool:
    tabs = [("alt", "h", "Inicio"), ("alt", "n", "Insertar"), ("alt", "g", "Diseño")]
    for alt_key, tab_key, tab_name in tabs:
        self.driver.send_key_combination(alt_key, tab_key)
        time.sleep(1)
        self.logger.info(f"✅ Navegando a {tab_name}")
```

## Técnicas Específicas de Word

### Navegación por Ribbon

#### Acceso a Pestañas
```python
# Método 1: Teclas de acceso
self.driver.send_key_combination("alt", "h")  # Inicio
self.driver.send_key_combination("alt", "n")  # Insertar
self.driver.send_key_combination("alt", "g")  # Diseño

# Método 2: Navegación directa
tab_element = self.driver.find_element_by_name("Inicio")
tab_element.click()
```

#### Búsqueda de Botones en Ribbon
```python
def find_ribbon_button(self, button_name: str):
    # Buscar en ribbon actual
    try:
        return self.driver.find_element_by_name(button_name)
    except:
        # Si no se encuentra, buscar en grupos específicos
        groups = self.driver.find_elements_by_class_name("NetUIRibbonGroup")
        for group in groups:
            try:
                return group.find_element_by_name(button_name)
            except:
                continue
    return None
```

### Manipulación de Texto

#### Inserción y Selección
```python
def insert_text(self, text: str):
    active_element = self.driver.driver.switch_to.active_element
    active_element.send_keys(text)

def select_all_text(self):
    self.driver.send_key_combination("ctrl", "a")

def select_line(self):
    self.driver.send_key_combination("home")
    self.driver.send_key_combination("shift", "end")
```

#### Formato de Texto
```python
def apply_bold(self):
    self.driver.send_key_combination("ctrl", "b")

def apply_italic(self):
    self.driver.send_key_combination("ctrl", "i")

def apply_underline(self):
    self.driver.send_key_combination("ctrl", "u")

def change_font_size(self, size: int):
    self.driver.send_key_combination("ctrl", "shift", "p")  # Seleccionar tamaño
    active_element = self.driver.driver.switch_to.active_element
    active_element.send_keys(str(size))
    self.driver.send_key_combination("enter")
```

### Operaciones con Tablas

#### Insertar Tabla
```python
def insert_table(self, rows: int, cols: int):
    # Ir a insertar
    self.driver.send_key_combination("alt", "n")
    time.sleep(1)
    
    # Seleccionar tabla
    self.driver.driver.switch_to.active_element.send_keys("t")
    time.sleep(1)
    
    # Insertar tabla específica
    self.driver.driver.switch_to.active_element.send_keys("i")
    time.sleep(1)
    
    # Configurar dimensiones
    active_element = self.driver.driver.switch_to.active_element
    # Navegar a campos de filas/columnas y configurar
```

#### Navegar en Tabla
```python
def move_to_next_cell(self):
    self.driver.send_key_combination("tab")

def move_to_previous_cell(self):
    self.driver.send_key_combination("shift", "tab")

def select_table(self):
    self.driver.send_key_combination("alt", "j", "l")  # Herramientas de tabla
    time.sleep(1)
    self.driver.send_key_combination("k")  # Seleccionar tabla
```

### Operaciones de Guardado

#### Guardar en Diferentes Formatos
```python
def save_as_format(self, filename: str, format_type: str):
    # Abrir diálogo guardar como
    self.driver.send_key_combination("ctrl", "shift", "s")
    time.sleep(2)
    
    # Cambiar tipo de archivo
    self.driver.send_key_combination("alt", "t")
    time.sleep(1)
    
    # Buscar formato específico
    active_element = self.driver.driver.switch_to.active_element
    active_element.send_keys(format_type)
    time.sleep(1)
    self.driver.send_key_combination("enter")
    
    # Escribir nombre de archivo
    self.driver.send_key_combination("alt", "n")
    active_element = self.driver.driver.switch_to.active_element
    active_element.send_keys(filename)
    
    # Confirmar guardado
    self.driver.send_key_combination("enter")
```

## Mejores Prácticas

### 1. Manejo de Tiempos de Espera

```python
# Esperas fijas solo cuando sea necesario
time.sleep(2)  # Para estabilización de UI

# Preferir esperas condicionales
def wait_for_element(self, locator, timeout=10):
    for _ in range(timeout):
        try:
            return self.driver.find_element(*locator)
        except:
            time.sleep(1)
    raise TimeoutException(f"Element not found: {locator}")
```

### 2. Captura de Evidencia

```python
def execute_step_with_evidence(self, step_name: str, action):
    try:
        self.logger.info(f"Ejecutando: {step_name}")
        result = action()
        self.driver.take_screenshot(f"{step_name}_success")
        return result
    except Exception as e:
        self.logger.error(f"Error en {step_name}: {e}")
        self.driver.take_screenshot(f"{step_name}_error")
        raise e
```

### 3. Verificación de Resultados

```python
def verify_text_formatted(self, expected_format: dict) -> bool:
    # Seleccionar texto para verificar formato
    self.driver.send_key_combination("ctrl", "a")
    
    # Verificar propiedades usando ribbon
    # Esto depende de la implementación específica
    return True  # Implementar lógica de verificación
```

### 4. Limpieza y Restauración

```python
def restore_word_state(self):
    try:
        # Cerrar documentos sin guardar
        while True:
            self.driver.send_key_combination("ctrl", "w")
            time.sleep(1)
            
            # Manejar diálogo de guardado
            try:
                # Presionar "No" si aparece diálogo
                self.driver.send_key_combination("alt", "n")
                time.sleep(1)
            except:
                break
                
    except Exception as e:
        self.logger.warning(f"Error en restauración: {e}")
```

## Solución de Problemas Específicos

### Problema: Elementos no encontrados

**Causas**:
- Word no completamente cargado
- Elemento fuera de vista
- Cambio de contexto (modal abierto)

**Soluciones**:
```python
# 1. Esperar estabilización
time.sleep(3)

# 2. Buscar en toda la jerarquía
def find_element_recursive(self, name: str):
    # Buscar en ventana principal
    try:
        return self.driver.find_element_by_name(name)
    except:
        pass
    
    # Buscar en modales
    try:
        modal = self.driver.find_element_by_class_name("NUIDialog")
        return modal.find_element_by_name(name)
    except:
        pass
    
    raise ElementNotFound(f"Element '{name}' not found")

# 3. Refrescar contexto
self.driver.send_key_combination("escape")  # Cerrar modales
self.driver.send_key_combination("alt")     # Refrescar ribbon
```

### Problema: Acciones no surten efecto

**Causas**:
- Elemento no tiene foco
- Acción bloqueada por modal
- Timing incorrecto

**Soluciones**:
```python
# 1. Asegurar foco
def ensure_focus(self, element):
    element.click()
    time.sleep(0.5)
    
# 2. Verificar estado antes de acción
def click_when_enabled(self, element):
    for _ in range(10):
        if element.is_enabled():
            element.click()
            return True
        time.sleep(0.5)
    return False

# 3. Usar métodos alternativos
def send_keys_alternative(self, keys: str):
    try:
        # Método 1: elemento activo
        active = self.driver.driver.switch_to.active_element
        active.send_keys(keys)
    except:
        # Método 2: combinación de teclas
        for key in keys:
            self.driver.send_key_combination(key)
```

### Problema: Diálogos inesperados

```python
def handle_unexpected_dialogs(self):
    common_dialogs = [
        ("Guardar cambios", "alt+n"),  # No guardar
        ("Error", "enter"),            # Aceptar error
        ("Advertencia", "enter")       # Aceptar advertencia
    ]
    
    for dialog_text, response in common_dialogs:
        try:
            dialog = self.driver.find_element_by_name(dialog_text)
            if dialog.is_displayed():
                self.driver.send_key_combination(*response.split("+"))
                return True
        except:
            continue
    return False
```

## Ejemplo Implementado: Operaciones Básicas

El proyecto incluye un ejemplo completo en `examples/word_examples/01_word_basic_operations.py` que demuestra:

### Funcionalidades Implementadas
1. **Inicialización segura de Word** con verificación de estado
2. **Navegación por ribbon** (Inicio, Insertar, Diseño)
3. **Verificación de elementos UI** críticos
4. **Manejo robusto de errores** con logging detallado
5. **Cierre controlado** con manejo de diálogos

### Estructura del Ejemplo
```python
class WordBasicOperationsExample:
    def setup_and_start(self) -> bool:
        # Configuración e inicio de Word
        
    def verify_word_interface(self) -> bool:
        # Verificación de interfaz principal
        
    def navigate_ribbon_tabs(self) -> bool:
        # Navegación por pestañas del ribbon
        
    def close_word(self) -> bool:
        # Cierre seguro de Word
```

## Extensión y Personalización

### Extender el Ejemplo Base

Para agregar más funcionalidades al ejemplo existente:

```python
# Agregar al archivo 01_word_basic_operations.py
def create_document(self) -> bool:
    """Crear un nuevo documento."""
    try:
        self.driver.send_key_combination("ctrl", "n")
        time.sleep(2)
        self.logger.info("✅ Documento creado")
        return True
    except Exception as e:
        self.logger.error(f"Error creando documento: {e}")
        return False

def insert_text(self, text: str) -> bool:
    """Insertar texto en el documento."""
    try:
        active_element = self.driver.driver.switch_to.active_element
        active_element.send_keys(text)
        self.logger.info(f"✅ Texto insertado: {text[:30]}...")
        return True
    except Exception as e:
        self.logger.error(f"Error insertando texto: {e}")
        return False
```

### Crear Nuevos Ejemplos

Usar el ejemplo base como plantilla para crear nuevas funcionalidades:

```python
# Nuevo archivo: 02_word_document_creation.py
from examples.word_examples.01_word_basic_operations import WordBasicOperationsExample

class WordDocumentCreationExample(WordBasicOperationsExample):
    def create_formatted_document(self) -> bool:
        """Crear documento con formato."""
        # Implementar lógica específica
        pass
```

## Casos de Uso Comunes

Basándose en el ejemplo implementado, estos son patrones típicos:

### 1. Verificación de Estado
```python
# Verificar que Word está listo
if "Word" in self.driver.get_current_window_title():
    self.logger.info("✅ Word verificado")
```

### 2. Navegación Segura
```python
# Navegar con verificación
self.driver.send_key_combination("alt", "h")  # Ir a Inicio
time.sleep(1)  # Esperar estabilización
```

### 3. Manejo de Errores
```python
# Patrón de try-catch con logging
try:
    # Operación
    self.logger.info("✅ Operación exitosa")
    return True
except Exception as e:
    self.logger.error(f"Error: {e}")
    self.driver.take_screenshot("error_context")
    return False
```

Esta guía se basa en el ejemplo práctico implementado. Para casos específicos, consultar el código en `01_word_basic_operations.py` y adaptar según necesidades particulares. logging, manejo de errores y verificación de estado.
"""Test simple de conexión a WinAppDriver"""
import json
from appium import webdriver

desired_caps = {
    "app": "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE",
    "platformName": "Windows",
    "deviceName": "WindowsPC"
}

print("Capabilities a enviar:")
print(json.dumps(desired_caps, indent=2))

try:
    driver = webdriver.Remote(
        command_executor="http://127.0.0.1:4723",
        desired_capabilities=desired_caps
    )
    print("✓ Conexión exitosa")
    driver.quit()
except Exception as e:
    print(f"✗ Error: {e}")

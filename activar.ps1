# Script de activacion del entorno virtual
Write-Host "Configurando entorno de automatizacion de Word..." -ForegroundColor Cyan

# Verificar Python
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "Python no instalado" -ForegroundColor Red
    Write-Host "Descargar desde: https://python.org" -ForegroundColor Yellow
    exit 1
}

$pythonVersion = python --version
Write-Host "Python detectado: $pythonVersion" -ForegroundColor Green

# Crear entorno virtual si no existe
if (-not (Test-Path "venv\Scripts\Activate.ps1")) {
    Write-Host "Creando entorno virtual..." -ForegroundColor Yellow
    python -m venv venv
}

# Activar entorno virtual
Write-Host "Activando entorno virtual..." -ForegroundColor Green
& "venv\Scripts\Activate.ps1"

# Actualizar pip
Write-Host "Actualizando pip..." -ForegroundColor Blue
python -m pip install --upgrade pip --quiet

# Instalar dependencias
Write-Host "Instalando dependencias..." -ForegroundColor Blue
pip install -r requirements.txt --quiet

# Configurar .env
if (-not (Test-Path ".env")) {
    if (Test-Path ".env.example") {
        Copy-Item ".env.example" ".env"
        Write-Host "Archivo .env creado - EDITAR con rutas correctas" -ForegroundColor Yellow
    }
}

# Crear directorios
$directories = @("reports", "reports\logs", "reports\screenshots")
foreach ($dir in $directories) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

# Verificar WinAppDriver
if (Get-Command WinAppDriver -ErrorAction SilentlyContinue) {
    Write-Host "WinAppDriver encontrado" -ForegroundColor Green
} else {
    Write-Host "WinAppDriver no encontrado" -ForegroundColor Yellow
    Write-Host "Descargar: https://github.com/Microsoft/WinAppDriver/releases" -ForegroundColor Yellow
}

# Verificar Word
$wordPaths = @(
    "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
)

$wordFound = $false
foreach ($path in $wordPaths) {
    if (Test-Path $path) {
        Write-Host "Word encontrado: $path" -ForegroundColor Green
        $wordFound = $true
        break
    }
}

if (-not $wordFound) {
    Write-Host "Word no encontrado en ubicaciones estandar" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "ENTORNO CONFIGURADO" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Proximos pasos:" -ForegroundColor White
Write-Host "1. Editar .env con rutas correctas" -ForegroundColor Gray
Write-Host "2. Iniciar WinAppDriver como Administrador" -ForegroundColor Gray
Write-Host "3. Ejecutar: python examples\word_examples\01_word_basic_operations.py" -ForegroundColor Gray
Write-Host ""

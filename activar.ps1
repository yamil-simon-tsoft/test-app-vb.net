# Script de activación del entorno virtual para PowerShell
# Uso: .\activar.ps1

Write-Host "🚀 Configurando entorno de automatización VB.NET..." -ForegroundColor Cyan

# Función para verificar si un comando existe
function Test-CommandExists {
    param($command)
    $oldPreference = $ErrorActionPreference
    $ErrorActionPreference = 'stop'
    try { 
        if(Get-Command $command) { 
            return $true 
        } 
    }
    catch { 
        return $false 
    }
    finally { 
        $ErrorActionPreference = $oldPreference 
    }
}

# Verificar Python
if (-not (Test-CommandExists "python")) {
    Write-Host "❌ Python no está instalado o no está en PATH" -ForegroundColor Red
    Write-Host "📥 Descargue Python desde: https://python.org" -ForegroundColor Yellow
    Read-Host "Presione Enter para continuar"
    exit 1
}

# Mostrar versión de Python
$pythonVersion = python --version
Write-Host "✅ $pythonVersion detectado" -ForegroundColor Green

# Verificar/Crear entorno virtual
if (-not (Test-Path "venv\Scripts\Activate.ps1")) {
    Write-Host "📦 Creando entorno virtual..." -ForegroundColor Yellow
    python -m venv venv
    if ($LASTEXITCODE -ne 0) {
        Write-Host "❌ Error al crear entorno virtual" -ForegroundColor Red
        exit 1
    }
}

# Activar entorno virtual
Write-Host "🔄 Activando entorno virtual..." -ForegroundColor Green
& "venv\Scripts\Activate.ps1"

if ($env:VIRTUAL_ENV) {
    Write-Host "✅ Entorno virtual activado: $env:VIRTUAL_ENV" -ForegroundColor Green
} else {
    Write-Host "⚠️ Advertencia: El entorno virtual no se activó correctamente" -ForegroundColor Yellow
}

# Actualizar pip
Write-Host "📥 Actualizando pip..." -ForegroundColor Blue
python -m pip install --upgrade pip --quiet

# Instalar dependencias
if (Test-Path "requirements.txt") {
    Write-Host "📦 Instalando dependencias..." -ForegroundColor Blue
    pip install -r requirements.txt --quiet
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ Dependencias instaladas correctamente" -ForegroundColor Green
    } else {
        Write-Host "⚠️ Algunas dependencias pueden no haberse instalado" -ForegroundColor Yellow
    }
} else {
    Write-Host "⚠️ No se encontró requirements.txt" -ForegroundColor Yellow
}

# Configurar archivo .env
if (-not (Test-Path ".env")) {
    if (Test-Path ".env.example") {
        Write-Host "📄 Creando archivo .env..." -ForegroundColor Blue
        Copy-Item ".env.example" ".env"
        Write-Host "⚙️ IMPORTANTE: Edite el archivo .env con las rutas correctas" -ForegroundColor Yellow
    }
}

# Crear directorios necesarios
$directories = @("reports", "reports\logs", "reports\screenshots", "reports\documents")
foreach ($dir in $directories) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Host "📁 Directorio creado: $dir" -ForegroundColor Gray
    }
}

# Verificar WinAppDriver
Write-Host "🔍 Verificando WinAppDriver..." -ForegroundColor Blue
if (Test-CommandExists "WinAppDriver") {
    Write-Host "✅ WinAppDriver encontrado" -ForegroundColor Green
} else {
    Write-Host "⚠️ WinAppDriver no encontrado" -ForegroundColor Yellow
    Write-Host "📥 Descárguelo desde: https://github.com/Microsoft/WinAppDriver/releases" -ForegroundColor Yellow
}

# Verificar Developer Mode
try {
    $regKey = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" -Name "AllowDevelopmentWithoutDevLicense" -ErrorAction SilentlyContinue
    if ($regKey -and $regKey.AllowDevelopmentWithoutDevLicense -eq 1) {
        Write-Host "✅ Developer Mode habilitado" -ForegroundColor Green
    } else {
        Write-Host "⚠️ Developer Mode no habilitado" -ForegroundColor Yellow
        Write-Host "⚙️ Active Developer Mode: Settings > Update & Security > For developers" -ForegroundColor Yellow
    }
} catch {
    Write-Host "⚠️ No se pudo verificar Developer Mode" -ForegroundColor Yellow
}

# Verificar Microsoft Word (opcional)
$wordPaths = @(
    "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
    "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
    "C:\Program Files\Microsoft Office\Office16\WINWORD.EXE"
)

$wordFound = $false
foreach ($path in $wordPaths) {
    if (Test-Path $path) {
        Write-Host "✅ Microsoft Word encontrado: $path" -ForegroundColor Green
        $wordFound = $true
        break
    }
}

if (-not $wordFound) {
    Write-Host "⚠️ Microsoft Word no encontrado en ubicaciones estándar" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "🎉 ENTORNO CONFIGURADO EXITOSAMENTE" -ForegroundColor Green
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "📋 Comandos disponibles:" -ForegroundColor White
Write-Host ""
Write-Host "📝 Ejemplos de Word:" -ForegroundColor Yellow
Write-Host "  python examples/word_examples/01_word_basic_operations.py" -ForegroundColor Gray
Write-Host "  python examples/word_examples/02_word_document_creation.py" -ForegroundColor Gray
Write-Host "  python examples/word_examples/03_word_text_formatting.py" -ForegroundColor Gray
Write-Host "  python examples/word_examples/04_word_table_operations.py" -ForegroundColor Gray
Write-Host "  python examples/word_examples/05_word_document_saving.py" -ForegroundColor Gray
Write-Host ""
Write-Host "🔧 Scripts de utilidad:" -ForegroundColor Yellow
Write-Host "  python scripts/run_word_examples.py          # Ejecutar todos los ejemplos" -ForegroundColor Gray
Write-Host "  python scripts/run_single_example.py --help  # Ayuda para ejemplo específico" -ForegroundColor Gray
Write-Host ""
Write-Host "📖 Documentación:" -ForegroundColor Yellow
Write-Host "  docs/installation.md     # Guía de instalación" -ForegroundColor Gray
Write-Host "  docs/usage.md           # Guía de uso" -ForegroundColor Gray
Write-Host "  docs/word_automation.md # Automatización Word" -ForegroundColor Gray
Write-Host ""
Write-Host "🔄 Para desactivar:" -ForegroundColor White
Write-Host "  deactivate" -ForegroundColor Gray
Write-Host ""

# Mostrar próximos pasos
Write-Host "🚀 PRÓXIMOS PASOS:" -ForegroundColor Cyan
Write-Host "1. Edite el archivo .env con las rutas correctas de sus aplicaciones" -ForegroundColor White
Write-Host "2. Inicie WinAppDriver como Administrador: WinAppDriver.exe" -ForegroundColor White
Write-Host "3. Ejecute un ejemplo: python examples/word_examples/01_word_basic_operations.py" -ForegroundColor White
Write-Host ""
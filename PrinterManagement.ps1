<#
.SYNOPSIS
    Advanced Printer Management Tool for Windows Environments.
.DESCRIPTION
    Automates printer installation, removal, test page dispatch, spooler cleanup,
    and inventorying with interactive TUI, GUI dialogs, logs, and progress reporting.
.VERSION
    3.2.0
#>

# Load Windows Forms for GUI dialogs if available
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
} catch {
    # Non-GUI environments (e.g. Server Core)
}

# Configuration
$script:LogFile = Join-Path -Path $(if ($PSScriptRoot) { $PSScriptRoot } else { "." }) -ChildPath "PrinterManagement.log"

# Localization
$script:Strings = @{
    'en' = @{
        'AdminError'        = "[!] ERROR: This script requires administrator privileges."
        'AdminPrompt'       = "Please right-click PowerShell and select 'Run as Administrator'."
        'PressEnterToExit'  = "Press Enter to exit..."
        'PressEnter'        = "`nPress Enter to return to menu..."
        'CatOperations'     = "OPERATIONS "
        'OptAddPrinters'    = "      [1]  Add Printers"
        'DescAddPrinters'   = "         Bulk CSV / TCP-IP / Shared UNC"
        'OptRemPrinters'    = "      [2]  Remove Printers"
        'DescRemPrinters'   = "      CSV or interactive selection"
        'CatDiagnostics'    = "DIAGNOSTICS "
        'OptSendTest'       = "      [3]  Send Test Pages"
        'DescSendTest'      = "     CIM dispatch with printui fallback"
        'OptClearQueue'     = "      [4]  Clear Print Queue"
        'DescClearQueue'    = "   Purge Spooler & restart service"
        'CatSystem'         = "SYSTEM "
        'OptInventory'      = "      [5]  Inventory Printers"
        'DescInventory'     = "  Console table & CSV export"
        'OptViewLog'        = "      [6]  View Activity Log"
        'DescViewLog'       = "   Review recent operations"
        'OptExit'           = "      [0]  Exit"
        'MadeWith'          = "                         Made with "
        'By'                = " by "
        'SelectOption'      = "   Select an option (0-6)"
        'Disconnecting'     = "   Disconnecting from PrinterManagement..."
        'SessionTerminated' = "   Session terminated. Goodbye!"
        'InvalidOption'     = "   Invalid option. Please choose a number between 0 and 6."
        'HeaderAdd'         = "`n=========================== [ 1. ADD PRINTERS ] ==========================="
        'HeaderRem'         = "`n========================= [ 2. REMOVE PRINTERS ] ========================="
        'HeaderTest'        = "`n======================= [ 3. SEND TEST PAGES ] ======================="
        'HeaderClear'       = "`n====================== [ 4. CLEAR PRINT QUEUES ] ======================"
        'HeaderInv'         = "`n===================== [ 5. INVENTORY PRINTERS ] ====================="
        'HeaderLog'         = "`n======================= [ 6. VIEW ACTIVITY LOG ] ======================="
        
        'InvLogStart'       = "Gathering system printer information..."
        'InvNoPrinters'     = "No installed printers found on this system."
        'InvExported'       = "Inventory exported to"
        'InvGridTitle'      = "Installed Printers Inventory"
        'InvErr'            = "Error generating printer inventory:"
        
        'RemOpt1'           = " [1] Remove printers specified in CSV"
        'RemOpt2'           = " [2] Select printers interactively from GUI list"
        'RemOpt3'           = " [3] Cancel"
        'RemSubOpt'         = "`nSelect an option (1-3)"
        'RemTitleCSV'       = "Select CSV File to Remove Printers"
        'RemGridTitle'      = "Select printers to remove (Hold Ctrl/Shift to select multiple)"
        'RemNoSel'          = "No printers selected."
        'RemErrGrid'        = "Error opening selection grid:"
        'RemNoSpec'         = "No printers specified for removal."
        'RemAtt'            = "`n[!] ATTENTION: You are about to remove"
        'RemConf'           = "Are you sure you want to proceed? (Y/N)"
        'RemAbo'            = "Printer removal aborted by user."
        'RemProgAct'        = "Removing Printers"
        'RemRem'            = "removed."
        'RemErrRem'         = "Error removing printer"
        'RemNotEx'          = "does not exist."
        'RemPort'           = "Port removed:"
        'RemErrPort'        = "Error removing port"
        'RemFin'            = "Removal process completed. Processed:"
        
        'AddTitleCSV'       = "Select CSV File to Add Printers"
        'AddStart'          = "Starting bulk printer installation..."
        'AddNoCSV'          = "No CSV file selected. Operation cancelled."
        'AddImp'            = "Importing printers from"
        'AddNoRec'          = "No valid printer records found in CSV."
        'AddProgAct'        = "Installing Printers"
        'AddSuccess'        = "Installed printer"
        'AddFail'           = "Failed to add printer"
        'AddFin'            = "Finished:"
        'AddInstSucc'       = "installed successfully,"
        'AddFailT'          = "failed."
        
        'TestTitleCSV'      = "Select CSV File for Test Pages"
        'TestStart'         = "Sending test pages to"
        'TestProgAct'       = "Dispatching Test Pages"
        'TestNotInst'       = "is not installed locally. Skipping."
        'TestDisp'          = "Test page dispatched to"
        'TestRetCode'       = "PrintTestPage returned code"
        'TestDispUI'        = "Test page dispatched via printui for"
        'TestFail'          = "Failed to send test page to"
        'TestFin'           = "Test pages dispatch process completed:"
        'TestDispT'         = "dispatched,"
        
        'ClearWarn'         = "This operation will stop the Spooler service and purge all pending jobs."
        'ClearConf'         = "`nAre you sure you want to stop Spooler and clear queues? (Y/N)"
        'ClearCanc'         = "Print queue clearing cancelled."
        'ClearStop'         = "Stopping Spooler service..."
        'ClearPurge'        = "Purging spool files from system directory..."
        'ClearStart'        = "Restarting Spooler service..."
        'ClearSucc'         = "Print queues cleared and Spooler restarted successfully!"
        'ClearErr'          = "Error while clearing print queues:"
        
        'LogLoc'            = "Log location:"
        'LogOpen'           = "`nOpen full log file in default editor? (Y/N)"
        'LogNo'             = "No activity log file found yet."
        
        'TmplSucc'          = "Template CSV created successfully at:"
        'TmplSamp'          = "`nSample content generated:"
        'TmplOpen'          = "`nDo you want to open the generated CSV file now? (Y/N)"
        'TmplErr'           = "Error creating template CSV:"
        
        'GridError'         = "Could not launch GridView:"
        'AddNoValid'        = "No valid printer records found in CSV."
        'AddStartInst'      = "Starting installation of {0} printer(s)..."
        'AddProgStat'       = "[{0}/{1}] Processing: {2}"
        'AddConnNet'        = "Connecting to network shared printer: '{0}'..."
        'AddConnSucc'       = "Successfully connected to shared printer '{0}'."
        'AddConnFail'       = "Failed to connect to shared printer '{0}': {1}"
        'AddSkipMiss'       = "Skipping '{0}': Missing Port or Driver."
        'AddPortSucc'       = "Created printer port '{0}'."
        'AddPortFail'       = "Error creating port '{0}': {1}"
        'AddDrvMiss'        = "Driver '{0}' is NOT installed on this system. Skipping '{1}'."
        'AddInstSucc2'      = "Installed printer '{0}' on port '{1}' with driver '{2}'."
        'AddInstFail2'      = "Failed to add printer '{0}': {1}"
        'AddFin2'           = "Finished: {0} installed successfully, {1} failed."
        
        'RemProgStat'       = "[{0}/{1}] Removing: {2}"
        'RemRemSucc'        = "Printer '{0}' removed."
        'RemRemFail'        = "Error removing printer '{0}': {1}"
        'RemRemMiss'        = "Printer '{0}' does not exist."
        'RemPortSucc'       = "Port '{0}' removed."
        'RemPortFail'       = "Error removing port '{0}': {1}"
        'RemFin2'           = "Removal process completed. Processed: {0} printer(s)."
        
        'TestNoValid'       = "No valid printers found in CSV."
        'TestStart2'        = "Sending test pages to {0} printer(s)..."
        'TestProgStat'      = "[{0}/{1}] Testing: {2}"
        'TestNotInst2'      = "Printer '{0}' is not installed locally. Skipping."
        'TestDispSucc'      = "Test page dispatched to '{0}'."
        'TestDispCode'      = "PrintTestPage returned code {0} for '{1}'."
        'TestDispUI2'       = "Test page dispatched via printui for '{0}'."
        'TestDispFail2'     = "Failed to send test page to '{0}': {1}"
        'TestFin2'          = "Test pages dispatch process completed: {0} dispatched, {1} failed."
        
        'ClearWarn2'        = "This operation will stop the Spooler service and purge all pending jobs."
        'ClearErr2'         = "Error while clearing print queues: {0}"
        
        'InvStart'          = "Gathering system printer information..."
        'InvNoPrint'        = "No installed printers found on this system."
        'InvExp2'           = "Inventory exported to '{0}' ({1} printer(s) found)."
        'InvGridErr'        = "Could not launch GridView: {0}"
        'InvErr2'           = "Error generating printer inventory: {0}"
        
        'TmplSucc2'         = "Template CSV created successfully at: {0}"
        'TmplFail2'         = "Error creating template CSV: {0}"
        
        'LogLoc2'           = "Log location: {0}`n"
        'LogNo2'            = "No activity log file found yet."
        
        'SelCSV'            = "`nSelect CSV File:"
        'SelCSVFail'        = "File does not exist or is not a file: {0}"

        'GetPathNoExist'    = "Specified path '{0}' does not exist."
        'GetPathGui'        = "  - Press [Enter] to open GUI File Explorer"
        'GetPathCli'        = "  - Or type / paste the file path directly"
        'GetPathPrompt'     = "File path"
        'GetPathGuiSel'     = "Selected file via GUI: {0}"
        'GetPathCancel'     = "File selection cancelled by user."
        'GetPathNoGui'      = "GUI file picker unavailable in this environment. Please enter path manually."
        'GetPathMan'        = "Enter CSV file path"
        'GetPathCliSel'     = "Selected file via CLI: {0}"
        'GetPathInv'        = "The path '{0}' is invalid or does not exist."
        
        'SmartNoFile'       = "File not found: {0}"
        'SmartRead'         = "Reading CSV with detected delimiter: '{0}'"
        'SmartErr'          = "Error parsing CSV file: {0}"
        'HeaderTmpl'        = "`n==================== [ 6. GENERATE CSV TEMPLATE ] ===================="

        'BannerSubtitle'    = "Windows Printer Management Suite"
        'InvColName'        = "Printer Name"
        'InvColPort'        = "Port"
        'InvColDriver'      = "Driver"
        'InvColDef'         = "Default"
        'InvYes'            = "YES"
        'InvNo'             = "No"
    }
    'es' = @{
        'AdminError'        = "[!] ERROR: Este script requiere privilegios de administrador."
        'AdminPrompt'       = "Por favor, haz clic derecho en PowerShell y selecciona 'Ejecutar como Administrador'."
        'PressEnterToExit'  = "Presiona Enter para salir..."
        'PressEnter'        = "`nPresiona Enter para volver al menu..."
        'CatOperations'     = "OPERACIONES "
        'OptAddPrinters'    = "      [1]  Agregar Impresoras"
        'DescAddPrinters'   = "         CSV Masivo / TCP-IP / Red UNC"
        'OptRemPrinters'    = "      [2]  Eliminar Impresoras"
        'DescRemPrinters'   = "      CSV o seleccion interactiva"
        'CatDiagnostics'    = "DIAGNOSTICO "
        'OptSendTest'       = "      [3]  Enviar Pagina de Prueba"
        'DescSendTest'      = "     Envio CIM con printui alternativo"
        'OptClearQueue'     = "      [4]  Limpiar Cola de Impresion"
        'DescClearQueue'    = "   Purgar Spooler y reiniciar servicio"
        'CatSystem'         = "SISTEMA "
        'OptInventory'      = "      [5]  Inventario de Impresoras"
        'DescInventory'     = "  Tabla en consola y exportacion CSV"
        'OptViewLog'        = "      [6]  Ver Registro de Actividad"
        'DescViewLog'       = "   Revisar las operaciones recientes"
        'OptExit'           = "      [0]  Salir"
        'MadeWith'          = "                         Hecho con "
        'By'                = " por "
        'SelectOption'      = "   Selecciona una opcion (0-6)"
        'Disconnecting'     = "   Desconectando de PrinterManagement..."
        'SessionTerminated' = "   Sesion terminada. Adios!"
        'InvalidOption'     = "   Opcion invalida. Por favor elige un numero entre 0 y 6."
        'HeaderAdd'         = "`n======================== [ 1. AGREGAR IMPRESORAS ] ========================"
        'HeaderRem'         = "`n======================= [ 2. ELIMINAR IMPRESORAS ] ========================"
        'HeaderTest'        = "`n====================== [ 3. ENVIAR PAGINA DE PRUEBA ] ====================="
        'HeaderClear'       = "`n====================== [ 4. LIMPIAR COLA DE IMPRESION ] ==================="
        'HeaderInv'         = "`n====================== [ 5. INVENTARIO DE IMPRESORAS ] ===================="
        'HeaderLog'         = "`n====================== [ 6. VER REGISTRO DE ACTIVIDAD ] ======================"
        
        'InvLogStart'       = "Recopilando informacion de impresoras del sistema..."
        'InvNoPrinters'     = "No se encontraron impresoras instaladas en este sistema."
        'InvExported'       = "Inventario exportado a"
        'InvGridTitle'      = "Inventario de Impresoras Instaladas"
        'InvErr'            = "Error generando inventario de impresoras:"
        
        'RemOpt1'           = " [1] Eliminar impresoras especificadas en CSV"
        'RemOpt2'           = " [2] Seleccionar impresoras interactivamente desde lista GUI"
        'RemOpt3'           = " [3] Cancelar"
        'RemSubOpt'         = "`nSelecciona una opcion (1-3)"
        'RemTitleCSV'       = "Selecciona Archivo CSV para Eliminar Impresoras"
        'RemGridTitle'      = "Selecciona impresoras a eliminar (Manten Ctrl/Shift para seleccion multiple)"
        'RemNoSel'          = "No se selecciono ninguna impresora."
        'RemErrGrid'        = "Error abriendo la tabla de seleccion:"
        'RemNoSpec'         = "No se especificaron impresoras para eliminar."
        'RemAtt'            = "`n[!] ATENCION: Estas a punto de eliminar"
        'RemConf'           = "Estas seguro de que deseas proceder? (S/N)"
        'RemAbo'            = "Eliminacion de impresoras abortada por el usuario."
        'RemProgAct'        = "Eliminando Impresoras"
        'RemRem'            = "eliminada."
        'RemErrRem'         = "Error eliminando impresora"
        'RemNotEx'          = "no existe."
        'RemPort'           = "Puerto eliminado:"
        'RemErrPort'        = "Error eliminando puerto"
        'RemFin'            = "Proceso de eliminacion completado. Procesadas:"
        
        'AddTitleCSV'       = "Selecciona Archivo CSV para Agregar Impresoras"
        'AddStart'          = "Iniciando instalacion masiva de impresoras..."
        'AddNoCSV'          = "No se selecciono un archivo CSV. Operacion cancelada."
        'AddImp'            = "Importando impresoras de"
        'AddNoRec'          = "No se encontraron registros de impresoras validos en el CSV."
        'AddProgAct'        = "Instalando Impresoras"
        'AddSuccess'        = "Impresora instalada"
        'AddFail'           = "Error al agregar la impresora"
        'AddFin'            = "Terminado:"
        'AddInstSucc'       = "instaladas exitosamente,"
        'AddFailT'          = "fallaron."
        
        'TestTitleCSV'      = "Selecciona Archivo CSV para Paginas de Prueba"
        'TestStart'         = "Enviando paginas de prueba a"
        'TestProgAct'       = "Enviando Paginas de Prueba"
        'TestNotInst'       = "no esta instalada localmente. Saltando."
        'TestDisp'          = "Pagina de prueba enviada a"
        'TestRetCode'       = "PrintTestPage devolvio codigo"
        'TestDispUI'        = "Pagina de prueba enviada via printui para"
        'TestFail'          = "Error enviando pagina de prueba a"
        'TestFin'           = "Proceso de envio de paginas de prueba completado:"
        'TestDispT'         = "enviadas,"
        
        'ClearWarn'         = "Esta operacion detendra el servicio Spooler y purgara todos los trabajos pendientes."
        'ClearConf'         = "`nEstas seguro que deseas detener el Spooler y limpiar las colas? (S/N)"
        'ClearCanc'         = "Limpieza de colas de impresion cancelada."
        'ClearStop'         = "Deteniendo servicio Spooler..."
        'ClearPurge'        = "Purgando archivos spool del directorio del sistema..."
        'ClearStart'        = "Reiniciando servicio Spooler..."
        'ClearSucc'         = "Colas de impresion limpiadas y Spooler reiniciado exitosamente!"
        'ClearErr'          = "Error al limpiar colas de impresion:"
        
        'LogLoc'            = "Ubicacion del registro:"
        'LogOpen'           = "`nAbrir el archivo de registro completo en el editor por defecto? (S/N)"
        'LogNo'             = "No se ha encontrado archivo de registro de actividad todavia."
        
        'TmplSucc'          = "CSV de plantilla creado exitosamente en:"
        'TmplSamp'          = "`nContenido de muestra generado:"
        'TmplOpen'          = "`nDeseas abrir el archivo CSV generado ahora? (S/N)"
        'TmplErr'           = "Error creando la plantilla CSV:"
        
        'GridError'         = "No se pudo lanzar GridView:"
        'AddNoValid'        = "No se encontraron registros de impresoras validos en el CSV."
        'AddStartInst'      = "Iniciando instalacion de {0} impresora(s)..."
        'AddProgStat'       = "[{0}/{1}] Procesando: {2}"
        'AddConnNet'        = "Conectando a impresora de red compartida: '{0}'..."
        'AddConnSucc'       = "Conectado exitosamente a la impresora compartida '{0}'."
        'AddConnFail'       = "Error al conectar a la impresora compartida '{0}': {1}"
        'AddSkipMiss'       = "Saltando '{0}': Falta Puerto o Controlador."
        'AddPortSucc'       = "Puerto de impresora '{0}' creado."
        'AddPortFail'       = "Error creando puerto '{0}': {1}"
        'AddDrvMiss'        = "El controlador '{0}' NO esta instalado. Saltando '{1}'."
        'AddInstSucc2'      = "Impresora '{0}' instalada en puerto '{1}' con controlador '{2}'."
        'AddInstFail2'      = "Error al agregar la impresora '{0}': {1}"
        'AddFin2'           = "Terminado: {0} instaladas exitosamente, {1} fallaron."
        
        'RemProgStat'       = "[{0}/{1}] Eliminando: {2}"
        'RemRemSucc'        = "Impresora '{0}' eliminada."
        'RemRemFail'        = "Error eliminando impresora '{0}': {1}"
        'RemRemMiss'        = "La impresora '{0}' no existe."
        'RemPortSucc'       = "Puerto '{0}' eliminado."
        'RemPortFail'       = "Error eliminando puerto '{0}': {1}"
        'RemFin2'           = "Proceso de eliminacion completado. Procesadas: {0} impresora(s)."
        
        'TestNoValid'       = "No se encontraron impresoras validas en el CSV."
        'TestStart2'        = "Enviando paginas de prueba a {0} impresora(s)..."
        'TestProgStat'      = "[{0}/{1}] Probando: {2}"
        'TestNotInst2'      = "Impresora '{0}' no instalada localmente. Saltando."
        'TestDispSucc'      = "Pagina de prueba enviada a '{0}'."
        'TestDispCode'      = "PrintTestPage devolvio codigo {0} para '{1}'."
        'TestDispUI2'       = "Pagina de prueba enviada via printui para '{0}'."
        'TestDispFail2'     = "Error al enviar pagina de prueba a '{0}': {1}"
        'TestFin2'          = "Envio de paginas de prueba completado: {0} enviadas, {1} fallaron."
        
        'ClearWarn2'        = "Esta operacion detendra el servicio Spooler y purgara todos los trabajos pendientes."
        'ClearErr2'         = "Error al limpiar colas de impresion: {0}"
        
        'InvStart'          = "Recopilando informacion de impresoras del sistema..."
        'InvNoPrint'        = "No se encontraron impresoras instaladas en este sistema."
        'InvExp2'           = "Inventario exportado a '{0}' ({1} impresora(s) encontradas)."
        'InvGridErr'        = "No se pudo lanzar GridView: {0}"
        'InvErr2'           = "Error generando inventario de impresoras: {0}"
        
        'TmplSucc2'         = "Plantilla CSV creada exitosamente en: {0}"
        'TmplFail2'         = "Error creando plantilla CSV: {0}"
        
        'LogLoc2'           = "Ubicacion del registro: {0}`n"
        'LogNo2'            = "No se encontro registro de actividad todavia."
        
        'SelCSV'            = "`nSelecciona un Archivo CSV:"
        'SelCSVFail'        = "El archivo no existe o no es valido: {0}"

        'GetPathNoExist'    = "La ruta especificada '{0}' no existe."
        'GetPathGui'        = "  - Presiona [Enter] para abrir el explorador de archivos GUI"
        'GetPathCli'        = "  - O escribe / pega la ruta del archivo directamente"
        'GetPathPrompt'     = "Ruta del archivo"
        'GetPathGuiSel'     = "Archivo seleccionado via GUI: {0}"
        'GetPathCancel'     = "Seleccion de archivo cancelada por el usuario."
        'GetPathNoGui'      = "Selector de archivos GUI no disponible. Introduce la ruta manualmente."
        'GetPathMan'        = "Introduce la ruta del archivo CSV"
        'GetPathCliSel'     = "Archivo seleccionado via CLI: {0}"
        'GetPathInv'        = "La ruta '{0}' es invalida o no existe."
        
        'SmartNoFile'       = "Archivo no encontrado: {0}"
        'SmartRead'         = "Leyendo CSV con delimitador detectado: '{0}'"
        'SmartErr'          = "Error al analizar el archivo CSV: {0}"
        'HeaderTmpl'        = "`n==================== [ 6. GENERAR PLANTILLA CSV ] ===================="

        'BannerSubtitle'    = "Suite de Administracion de Impresoras Windows"
        'InvColName'        = "Nombre Impresora"
        'InvColPort'        = "Puerto"
        'InvColDriver'      = "Controlador"
        'InvColDef'         = "Predet."
        'InvYes'            = "SI"
        'InvNo'             = "No"
    }
}

$lang = (Get-UICulture).TwoLetterISOLanguageName
if (-not $script:Strings.ContainsKey($lang)) { $lang = 'en' }
$script:T = $script:Strings[$lang]

# Logging function
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [ValidateSet("INFO", "SUCCESS", "WARN", "ERROR")]
        [string]$Level = "INFO",

        [string]$CustomLogPath = $script:LogFile
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $logEntry = "[$timestamp] [$Level] $Message"

    $colorMap = @{
        "INFO"    = "Cyan"
        "SUCCESS" = "Green"
        "WARN"    = "Yellow"
        "ERROR"   = "Red"
    }

    $iconMap = @{
        "INFO"    = "[i]"
        "SUCCESS" = "[OK]"
        "WARN"    = "[!]"
        "ERROR"   = "[FAIL]"
    }

    Write-Host "$($iconMap[$Level]) $Message" -ForegroundColor $colorMap[$Level]

    # Append to log file
    try {
        Add-Content -Path $CustomLogPath -Value $logEntry -Encoding UTF8 -ErrorAction SilentlyContinue
    } catch {
        # Silent continue if log file write fails
    }
}

# Function to display banner
function Show-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ===========================================================================" -ForegroundColor DarkCyan
    Write-Host "    ___      _      __          __  ___                                      " -ForegroundColor Cyan
    Write-Host "   / _ \____(_)__  / /____ ____/  |/  /__ ____  ___ ____ ____ __ _  ___ ___  " -ForegroundColor Cyan
    Write-Host "  / ___/ __/ / _ \/ __/ -_) __/ /|_/ / _ ``/ _ \/ _ ``/ _ ``/ -_)  ' \/ -_) _ \ " -ForegroundColor Cyan
    Write-Host " /_/  /_/ /_/_//_/\__/\__/_/ /_/  /_/\_,_/_//_/\_,_/\_, /\__/_/_/_/\__/_//_/ " -ForegroundColor Cyan
    Write-Host "                                                    /___/                    " -ForegroundColor Cyan
    Write-Host "  ===========================================================================" -ForegroundColor DarkCyan
    Write-Host "   v3.2.0" -ForegroundColor DarkGray -NoNewline
    Write-Host " | " -ForegroundColor DarkCyan -NoNewline
    Write-Host $script:T.BannerSubtitle -ForegroundColor White
    Write-Host "   github.com/IamCarron/PrinterManagement" -ForegroundColor DarkGray
    Write-Host "  ===========================================================================" -ForegroundColor DarkCyan
    Write-Host ""
}

# GUI & CLI File Picker
function Get-ValidFilePath {
    param (
        [string]$FilePath = "",
        [string]$Title = "Select Printer CSV File",
        [string]$Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    )

    if (-not [string]::IsNullOrWhiteSpace($FilePath)) {
        $clean = $FilePath.Trim('"').Trim("'")
        if (Test-Path -Path $clean -PathType Leaf) {
            return (Resolve-Path -Path $clean).Path
        }
        Write-Log ($script:T.GetPathNoExist -f $clean) "ERROR"
        return $null
    }

    Write-Host $script:T.SelCSV -ForegroundColor Yellow
    Write-Host $script:T.GetPathGui -ForegroundColor DarkGray
    Write-Host $script:T.GetPathCli -ForegroundColor DarkGray
    $inputPath = Read-Host -Prompt $script:T.GetPathPrompt

    if ([string]::IsNullOrWhiteSpace($inputPath)) {
        # Attempt to open GUI File Dialog
        try {
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.InitialDirectory = [System.IO.Directory]::GetCurrentDirectory()
            $openFileDialog.Filter = $Filter
            $openFileDialog.Title = $Title
            $openFileDialog.ShowHelp = $false

            $dialogResult = $openFileDialog.ShowDialog()
            if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
                Write-Log ($script:T.GetPathGuiSel -f $openFileDialog.FileName) "INFO"
                return $openFileDialog.FileName
            } else {
                Write-Log $script:T.GetPathCancel "WARN"
                return $null
            }
        } catch {
            Write-Log $script:T.GetPathNoGui "WARN"
            $manualPath = Read-Host -Prompt $script:T.GetPathMan
            if (-not [string]::IsNullOrWhiteSpace($manualPath) -and (Test-Path -Path $manualPath.Trim('"').Trim("'") -PathType Leaf)) {
                return (Resolve-Path -Path $manualPath.Trim('"').Trim("'")).Path
            }
            return $null
        }
    } else {
        $cleanPath = $inputPath.Trim('"').Trim("'")
        if (Test-Path -Path $cleanPath -PathType Leaf) {
            $resolved = (Resolve-Path -Path $cleanPath).Path
            Write-Log ($script:T.GetPathCliSel -f $resolved) "INFO"
            return $resolved
        } else {
            Write-Log ($script:T.GetPathInv -f $cleanPath) "ERROR"
            return $null
        }
    }
}

# Smart CSV Importer
function Import-SmartCsv {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -Path $Path -PathType Leaf)) {
        Write-Log ($script:T.SmartNoFile -f $Path) "ERROR"
        return $null
    }

    # Detect delimiter safely
    $firstLine = (Get-Content -Path $Path -TotalCount 1)
    $delimiter = ';'
    if ($firstLine -match ';') {
        $delimiter = ';'
    } elseif ($firstLine -match ',') {
        $delimiter = ','
    } elseif ($firstLine -match "`t") {
        $delimiter = "`t"
    }

    Write-Log ($script:T.SmartRead -f $delimiter) "INFO"

    try {
        $data = Import-Csv -Path $Path -Delimiter $delimiter -Encoding UTF8 -ErrorAction Stop

        # Normalize properties
        $normalizedList = @()
        foreach ($row in $data) {
            $nameVal = if ($row.Name) { $row.Name } elseif ($row.PrinterName) { $row.PrinterName } else { "" }
            $portVal = if ($row.LocalPort) { $row.LocalPort } elseif ($row.Port) { $row.Port } elseif ($row.PortName) { $row.PortName } else { "" }
            $driverVal = if ($row.DriverName) { $row.DriverName } elseif ($row.Driver) { $row.Driver } else { "" }

            if (-not [string]::IsNullOrWhiteSpace($nameVal)) {
                $normalizedList += [PSCustomObject]@{
                    Name       = $nameVal.ToString().Trim()
                    LocalPort  = $portVal.ToString().Trim()
                    DriverName = $driverVal.ToString().Trim()
                }
            }
        }
        return $normalizedList
    } catch {
        Write-Log ($script:T.SmartErr -f $_) "ERROR"
        return $null
    }
}

# Function to add printers
function Add-Printers {
    param (
        [string]$FilePath = ""
    )

    Write-Host $script:T.HeaderAdd -ForegroundColor Yellow

    $printersFile = Get-ValidFilePath -FilePath $FilePath -Title $script:T.AddTitleCSV
    if (-not $printersFile) {
        if (-not $FilePath) { Read-Host -Prompt $script:T.PressEnter }
        return @{ Total = 0; Success = 0; Failed = 0 }
    }

    $printerList = @(Import-SmartCsv -Path $printersFile)
    if (-not $printerList -or $printerList.Count -eq 0) {
        Write-Log $script:T.AddNoValid "WARN"
        if (-not $FilePath) { Read-Host -Prompt $script:T.PressEnter }
        return @{ Total = 0; Success = 0; Failed = 0 }
    }

    $total = $printerList.Count
    Write-Log ($script:T.AddStartInst -f $total) "INFO"

    $index = 0
    $successCount = 0
    $failCount = 0

    foreach ($printer in $printerList) {
        $index++
        $percent = [int](($index / $total) * 100)
        Write-Progress -Activity $script:T.AddProgAct -Status ($script:T.AddProgStat -f $index, $total, $printer.Name) -PercentComplete $percent

        $pName   = $printer.Name
        $pPort   = $printer.LocalPort
        $pDriver = $printer.DriverName

        # Case 1: Shared Network Printer (UNC Path: \\server\printer)
        if ($pPort -like "\\*" -or $pName -like "\\*") {
            $connectionPath = if ($pPort -like "\\*") { $pPort } else { $pName }
            Write-Log ($script:T.AddConnNet -f $connectionPath) "INFO"
            try {
                Add-Printer -ConnectionName $connectionPath -ErrorAction Stop
                Write-Log ($script:T.AddConnSucc -f $connectionPath) "SUCCESS"
                $successCount++
            } catch {
                Write-Log ($script:T.AddConnFail -f $connectionPath, $_) "ERROR"
                $failCount++
            }
            continue
        }

        # Case 2: Standard Printer (Local / TCP-IP)
        if ([string]::IsNullOrWhiteSpace($pPort) -or [string]::IsNullOrWhiteSpace($pDriver)) {
            Write-Log ($script:T.AddSkipMiss -f $pName) "WARN"
            $failCount++
            continue
        }

        # Create Port if needed
        $portExists = Get-PrinterPort -Name $pPort -ErrorAction SilentlyContinue
        if (-not $portExists) {
            try {
                # If IPv4 address or standard TCP hostname
                $isIpAddress = $pPort -match "^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$"
                $isHostname  = ($pPort -match "^[a-zA-Z0-9\.\-_]+$" -and $pPort -notmatch "^(USB|COM|LPT|FILE|PORTPROMPT|NUL)\d*[:]?$")

                if ($isIpAddress -or $isHostname) {
                    Add-PrinterPort -Name $pPort -PrinterHostAddress $pPort -ErrorAction Stop
                } else {
                    Add-PrinterPort -Name $pPort -ErrorAction Stop
                }
                Write-Log ($script:T.AddPortSucc -f $pPort) "SUCCESS"
            } catch {
                Write-Log ($script:T.AddPortFail -f $pPort, $_) "ERROR"
                $failCount++
                continue
            }
        }

        # Verify Driver
        $driverExists = Get-PrinterDriver -Name $pDriver -ErrorAction SilentlyContinue
        if (-not $driverExists) {
            Write-Log ($script:T.AddDrvMiss -f $pDriver, $pName) "ERROR"
            $failCount++
            continue
        }

        # Add Printer
        try {
            Add-Printer -Name $pName -PortName $pPort -DriverName $pDriver -ErrorAction Stop
            Write-Log ($script:T.AddInstSucc2 -f $pName, $pPort, $pDriver) "SUCCESS"
            $successCount++
        } catch {
            Write-Log ($script:T.AddInstFail2 -f $pName, $_) "ERROR"
            $failCount++
        }
    }

    Write-Progress -Activity $script:T.AddProgAct -Completed
    Write-Host "`n---------------------------------------------------------" -ForegroundColor DarkGray
    Write-Log ($script:T.AddFin2 -f $successCount, $failCount) $(if ($failCount -eq 0) { "SUCCESS" } else { "WARN" })
    if (-not $FilePath) { Read-Host -Prompt $script:T.PressEnter }

    return @{ Total = $total; Success = $successCount; Failed = $failCount }
}

# Function to remove printers
function Remove-Printers {
    param (
        [string]$FilePath = "",
        [array]$PrinterList = @(),
        [switch]$Force
    )

    Write-Host $script:T.HeaderRem -ForegroundColor Yellow

    $printersToRemove = @()

    if ($PrinterList -and $PrinterList.Count -gt 0) {
        $printersToRemove = $PrinterList
    } elseif (-not [string]::IsNullOrWhiteSpace($FilePath)) {
        $printersToRemove = @(Import-SmartCsv -Path $FilePath)
    } else {
        Write-Host $script:T.RemOpt1 -ForegroundColor Cyan
        Write-Host $script:T.RemOpt2 -ForegroundColor Cyan
        Write-Host $script:T.RemOpt3 -ForegroundColor DarkGray
        $subOption = Read-Host $script:T.RemSubOpt

        if ($subOption -eq "1") {
            $printersFile = Get-ValidFilePath -Title $script:T.RemTitleCSV
            if (-not $printersFile) { return @{ Total = 0; Success = 0; Failed = 0 } }
            $printersToRemove = @(Import-SmartCsv -Path $printersFile)
        } elseif ($subOption -eq "2") {
            try {
                $installed = Get-CimInstance -ClassName Win32_Printer | Select-Object Name, PortName, DriverName
                $selected = $installed | Out-GridView -Title $script:T.RemGridTitle -PassThru
                if ($selected) {
                    $printersToRemove = @($selected | ForEach-Object {
                        [PSCustomObject]@{
                            Name      = $_.Name
                            LocalPort = $_.PortName
                        }
                    })
                } else {
                    Write-Log $script:T.RemNoSel "WARN"
                    Read-Host -Prompt $script:T.PressEnter
                    return @{ Total = 0; Success = 0; Failed = 0 }
                }
            } catch {
                Write-Log "$($script:T.RemErrGrid) $_" "ERROR"
                return @{ Total = 0; Success = 0; Failed = 0 }
            }
        } else {
            return @{ Total = 0; Success = 0; Failed = 0 }
        }
    }

    if (-not $printersToRemove -or $printersToRemove.Count -eq 0) {
        Write-Log $script:T.RemNoSpec "WARN"
        if (-not $Force) { Read-Host -Prompt $script:T.PressEnter }
        return @{ Total = 0; Success = 0; Failed = 0 }
    }

    # Security Confirmation
    if (-not $Force) {
        Write-Host "$($script:T.RemAtt) $($printersToRemove.Count)." -ForegroundColor Yellow
        $confirm = Read-Host $script:T.RemConf
        if ($confirm -notmatch '^(y|s|yes|si)$') {
            Write-Log $script:T.RemAbo "INFO"
            Read-Host -Prompt $script:T.PressEnter
            return @{ Total = 0; Success = 0; Failed = 0 }
        }
    }

    $total = $printersToRemove.Count
    $index = 0
    $successCount = 0
    $failCount = 0

    foreach ($item in $printersToRemove) {
        $index++
        $pName = $item.Name
        $pPort = if ($item.LocalPort) { $item.LocalPort } else { "" }

        Write-Progress -Activity $script:T.RemProgAct -Status ($script:T.RemProgStat -f $index, $total, $pName) -PercentComplete $([int](($index / $total) * 100))

        # Remove Printer
        $printerExists = Get-Printer -Name $pName -ErrorAction SilentlyContinue
        if ($printerExists) {
            try {
                Remove-Printer -Name $pName -ErrorAction Stop
                Write-Log ($script:T.RemRemSucc -f $pName) "SUCCESS"
                $successCount++
            } catch {
                Write-Log ($script:T.RemRemFail -f $pName, $_) "ERROR"
                $failCount++
            }
        } else {
            Write-Log ($script:T.RemRemMiss -f $pName) "WARN"
        }

        # Remove Port if provided and port exists
        if (-not [string]::IsNullOrWhiteSpace($pPort)) {
            $portExists = Get-PrinterPort -Name $pPort -ErrorAction SilentlyContinue
            if ($portExists) {
                try {
                    Remove-PrinterPort -Name $pPort -ErrorAction Stop
                    Write-Log ($script:T.RemPortSucc -f $pPort) "SUCCESS"
                } catch {
                    Write-Log ($script:T.RemPortFail -f $pPort, $_) "WARN"
                }
            }
        }
    }

    Write-Progress -Activity $script:T.RemProgAct -Completed
    Write-Host "`n---------------------------------------------------------" -ForegroundColor DarkGray
    Write-Log ($script:T.RemFin2 -f $total) "SUCCESS"
    if (-not $Force) { Read-Host -Prompt $script:T.PressEnter }

    return @{ Total = $total; Success = $successCount; Failed = $failCount }
}

# Function to send test pages
function Send-TestPages {
    param (
        [string]$FilePath = "",
        [array]$PrinterList = @()
    )

    Write-Host $script:T.HeaderTest -ForegroundColor Yellow

    $printersToTest = @()
    if ($PrinterList -and $PrinterList.Count -gt 0) {
        $printersToTest = $PrinterList
    } elseif (-not [string]::IsNullOrWhiteSpace($FilePath)) {
        $printersToTest = @(Import-SmartCsv -Path $FilePath)
    } else {
        $printersFile = Get-ValidFilePath -Title "Select CSV File for Test Pages"
        if (-not $printersFile) {
            Read-Host -Prompt $script:T.PressEnter
            return @{ Total = 0; Success = 0; Failed = 0 }
        }
        $printersToTest = @(Import-SmartCsv -Path $printersFile)
    }

    if (-not $printersToTest -or $printersToTest.Count -eq 0) {
        Write-Log $script:T.TestNoValid "WARN"
        if (-not $FilePath) { Read-Host -Prompt $script:T.PressEnter }
        return @{ Total = 0; Success = 0; Failed = 0 }
    }

    $total = $printersToTest.Count
    $index = 0
    $successCount = 0
    $failCount = 0

    Write-Log ($script:T.TestStart2 -f $total) "INFO"

    foreach ($printer in $printersToTest) {
        $index++
        $pName = $printer.Name
        Write-Progress -Activity $script:T.TestProgAct -Status ($script:T.TestProgStat -f $index, $total, $pName) -PercentComplete $([int](($index / $total) * 100))

        $printerObj = Get-Printer -Name $pName -ErrorAction SilentlyContinue
        if (-not $printerObj) {
            Write-Log ($script:T.TestNotInst2 -f $pName) "WARN"
            $failCount++
            continue
        }

        # Escape single quotes in printer name for WQL filter
        $escapedName = $pName -replace "'", "''"

        try {
            $wmiPrinter = Get-CimInstance -ClassName Win32_Printer -Filter "Name='$escapedName'" -ErrorAction Stop
            $result = Invoke-CimMethod -InputObject $wmiPrinter -MethodName "PrintTestPage" -ErrorAction Stop
            if ($result.ReturnValue -eq 0) {
                Write-Log ($script:T.TestDispSucc -f $pName) "SUCCESS"
                $successCount++
            } else {
                Write-Log ($script:T.TestDispCode -f $result.ReturnValue, $pName) "WARN"
                $failCount++
            }
        } catch {
            try {
                Start-Process rundll32.exe -ArgumentList "printui.dll,PrintUIEntry /k /n `"$pName`"" -NoNewWindow -Wait
                Write-Log ($script:T.TestDispUI2 -f $pName) "SUCCESS"
                $successCount++
            } catch {
                Write-Log ($script:T.TestDispFail2 -f $pName, $_) "ERROR"
                $failCount++
            }
        }
    }

    Write-Progress -Activity $script:T.TestProgAct -Completed
    Write-Host "`n---------------------------------------------------------" -ForegroundColor DarkGray
    Write-Log ($script:T.TestFin2 -f $successCount, $failCount) "SUCCESS"
    if (-not $FilePath) { Read-Host -Prompt $script:T.PressEnter }

    return @{ Total = $total; Success = $successCount; Failed = $failCount }
}

# Function to clear print queues
function Clear-PrintQueues {
    param (
        [switch]$Force
    )

    Write-Host $script:T.HeaderClear -ForegroundColor Yellow
    Write-Log $script:T.ClearWarn2 "WARN"

    if (-not $Force) {
        $confirm = Read-Host $script:T.ClearConf
        if ($confirm -notmatch '^(y|s|yes|si)$') {
            Write-Log $script:T.ClearCanc "INFO"
            Read-Host -Prompt $script:T.PressEnter
            return $false
        }
    }

    try {
        Write-Log $script:T.ClearStop "INFO"
        Stop-Service -Name Spooler -Force -ErrorAction Stop

        $spooler = Get-Service -Name Spooler
        $timeout = 10
        while ($spooler.Status -ne 'Stopped' -and $timeout -gt 0) {
            Start-Sleep -Milliseconds 500
            $spooler.Refresh()
            $timeout--
        }

        # Extra 1s wait to ensure file handles release
        Start-Sleep -Seconds 1

        Write-Log $script:T.ClearPurge "INFO"
        Remove-Item -Path "$env:windir\System32\spool\PRINTERS\*" -Force -Recurse -ErrorAction SilentlyContinue

        Write-Log $script:T.ClearStart "INFO"
        Start-Service -Name Spooler -ErrorAction Stop

        Write-Log $script:T.ClearSucc "SUCCESS"
        if (-not $Force) { Read-Host -Prompt $script:T.PressEnter }
        return $true
    } catch {
        Write-Log ($script:T.ClearErr2 -f $_) "ERROR"
        if (-not $Force) { Read-Host -Prompt $script:T.PressEnter }
        return $false
    }
}

# Function to inventory printers
function Inventory-Printers {
    param (
        [string]$OutputPath = ".\inventory.csv",
        [switch]$NoTable,
        [switch]$NoPrompt
    )

    Write-Host $script:T.HeaderInv -ForegroundColor Yellow

    try {
        Write-Log $script:T.InvLogStart "INFO"
        $printers = @(Get-CimInstance -ClassName Win32_Printer | Select-Object Name, DriverName, PortName, ShareName, Published, PrinterStatus, Default, Location, Comment)

        if (-not $printers -or $printers.Count -eq 0) {
            Write-Log $script:T.InvNoPrinters "WARN"
            if (-not $NoTable -and -not $NoPrompt) { Read-Host -Prompt $script:T.PressEnter }
            return @()
        }

        # Export to CSV
        $printers | Export-Csv -Path $OutputPath -Delimiter ';' -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        Write-Log "$($script:T.InvExported) '$OutputPath' ($($printers.Count))." "SUCCESS"

        if (-not $NoTable) {
            Write-Host "`n$($script:T.InvGridTitle):" -ForegroundColor Cyan
            $displayData = $printers | Select-Object @{Name=$script:T.InvColName; Expression={$_.Name}},
                                                     @{Name=$script:T.InvColPort; Expression={$_.PortName}},
                                                     @{Name=$script:T.InvColDriver; Expression={$_.DriverName}},
                                                     @{Name=$script:T.InvColDef; Expression={if ($_.Default) { $script:T.InvYes } else { $script:T.InvNo }}}
            $displayData | Format-Table -AutoSize | Out-String | Write-Host
        }

        return $printers
    } catch {
        Write-Log "$($script:T.InvErr) $_" "ERROR"
        return $null
    } finally {
        if (-not $NoTable -and -not $NoPrompt) { Read-Host -Prompt $script:T.PressEnter }
    }
}

# Function to generate CSV template
function New-PrinterTemplateCsv {
    param (
        [string]$OutputPath = ".\template_printers.csv",
        [switch]$NoOpen
    )

    Write-Host $script:T.HeaderTmpl -ForegroundColor Yellow

    $sampleData = @"
Name;LocalPort;DriverName
Office_HP_LaserJet;192.168.1.50;HP Universal Printing PCL 6
Finance_Canon_MFP;\\printserver01\Canon_Finance;Canon Generic Plus PCL6
Warehouse_Zebra_Labels;USB001;ZDesigner ZD420-203dpi ZPL
HR_Epson_WorkForce;192.168.1.55;EPSON WF-C5790 Series
"@

    try {
        Set-Content -Path $OutputPath -Value $sampleData -Encoding UTF8 -ErrorAction Stop
        Write-Log ($script:T.TmplSucc2 -f $OutputPath) "SUCCESS"
        Write-Host $script:T.TmplSamp -ForegroundColor Cyan
        Write-Host $sampleData -ForegroundColor DarkGray

        if (-not $NoOpen) {
            $open = Read-Host $script:T.TmplOpen
            if ($open -match '^(y|s|yes|si)$') {
                try { Invoke-Item $OutputPath } catch { }
            }
        }
        return $OutputPath
    } catch {
        Write-Log ($script:T.TmplFail2 -f $_) "ERROR"
        return $null
    } finally {
        if (-not $NoOpen) { Read-Host -Prompt $script:T.PressEnter }
    }
}

# Function to view activity log
function Show-ActivityLog {
    param (
        [string]$LogPath = $script:LogFile,
        [int]$Tail = 25
    )

    Write-Host $script:T.HeaderLog -ForegroundColor Yellow

    if (Test-Path -Path $LogPath) {
        Write-Host ($script:T.LogLoc2 -f $LogPath) -ForegroundColor DarkGray
        Get-Content -Path $LogPath -Tail $Tail | ForEach-Object {
            if ($_ -match "\[SUCCESS\]") { Write-Host $_ -ForegroundColor Green }
            elseif ($_ -match "\[ERROR\]") { Write-Host $_ -ForegroundColor Red }
            elseif ($_ -match "\[WARN\]") { Write-Host $_ -ForegroundColor Yellow }
            else { Write-Host $_ -ForegroundColor Cyan }
        }

        $open = Read-Host $script:T.LogOpen
        if ($open -match '^(y|s|yes|si)$') {
            try { Invoke-Item $LogPath } catch { }
        }
    } else {
        Write-Log $script:T.LogNo2 "INFO"
        Read-Host -Prompt $script:T.PressEnter
    }
}

# Function to start the interactive management console
function Start-PrinterManagement {
    [CmdletBinding()]
    param()

    # Check for Administrator privileges
    $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    if (-not $isAdmin) {
        Write-Host "`n$($script:T.AdminError)" -ForegroundColor Red
        Write-Host "$($script:T.AdminPrompt)`n" -ForegroundColor Yellow
        Read-Host -Prompt $script:T.PressEnterToExit
        return
    }

    do {
        Show-Banner

        Write-Host "   >> " -ForegroundColor Green -NoNewline
        Write-Host $($script:T.CatOperations) -ForegroundColor Green -NoNewline
        Write-Host "-------------------------------------------------" -ForegroundColor DarkGray
        Write-Host $($script:T.OptAddPrinters) -ForegroundColor Cyan -NoNewline
        Write-Host $($script:T.DescAddPrinters) -ForegroundColor DarkGray
        Write-Host $($script:T.OptRemPrinters) -ForegroundColor Cyan -NoNewline
        Write-Host $($script:T.DescRemPrinters) -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "   >> " -ForegroundColor Green -NoNewline
        Write-Host $($script:T.CatDiagnostics) -ForegroundColor Green -NoNewline
        Write-Host "------------------------------------------------" -ForegroundColor DarkGray
        Write-Host $($script:T.OptSendTest) -ForegroundColor Cyan -NoNewline
        Write-Host $($script:T.DescSendTest) -ForegroundColor DarkGray
        Write-Host $($script:T.OptClearQueue) -ForegroundColor Cyan -NoNewline
        Write-Host $($script:T.DescClearQueue) -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "   >> " -ForegroundColor Green -NoNewline
        Write-Host $($script:T.CatSystem) -ForegroundColor Green -NoNewline
        Write-Host "-----------------------------------------------------" -ForegroundColor DarkGray
        Write-Host $($script:T.OptInventory) -ForegroundColor Cyan -NoNewline
        Write-Host $($script:T.DescInventory) -ForegroundColor DarkGray
        Write-Host $($script:T.OptViewLog) -ForegroundColor Cyan -NoNewline
        Write-Host $($script:T.DescViewLog) -ForegroundColor DarkGray
        Write-Host ""
        Write-Host "   ---------------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host $($script:T.OptExit) -ForegroundColor Yellow
        Write-Host ""
        Write-Host $($script:T.MadeWith) -ForegroundColor DarkGray -NoNewline
        Write-Host "<3" -ForegroundColor Red -NoNewline
        Write-Host $($script:T.By) -ForegroundColor DarkGray -NoNewline
        Write-Host "IamCarron" -ForegroundColor Magenta
        Write-Host ""

        $option = Read-Host $($script:T.SelectOption)

        switch ($option) {
            "1" { $null = Add-Printers }
            "2" { $null = Remove-Printers }
            "3" { $null = Send-TestPages }
            "4" { $null = Clear-PrintQueues }
            "5" { $null = Inventory-Printers }
            "6" { $null = Show-ActivityLog }
            "0" {
                Write-Host ""
                Write-Host $($script:T.Disconnecting) -ForegroundColor DarkCyan
                Write-Host "$($script:T.SessionTerminated)`n" -ForegroundColor Cyan
            }
            default {
                Write-Warning $($script:T.InvalidOption)
                Start-Sleep -Seconds 1
            }
        }
    } while ($option -ne "0")
}

# Auto-run menu when executed directly (not in test mode)
if (-not $env:PRINTER_MANAGEMENT_TEST_MODE) {
    Start-PrinterManagement
}


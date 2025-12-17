# ========================================
# Ejecutor de Consultas SQL - Oracle
# Oracle SQLcl con exportacion a CSV y Excel
# PowerShell Script usando Microsoft Excel
# ========================================

# Configurar codificacion UTF-8 para la consola
$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::InputEncoding = [System.Text.Encoding]::UTF8

try {
    # Intentar cambiar la pagina de codigos a UTF-8
    $null = chcp 65001 2>$null
} catch {
    # Si falla, continuar sin chcp
}

# Limpiar pantalla para refrescar la codificacion
Clear-Host

# Funcion para verificar si Excel esta instalado
function Test-ExcelInstalled {
    try {
        # Intentar crear objeto Excel
        $excel = New-Object -ComObject "Excel.Application" -ErrorAction Stop
        $version = $excel.Version
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        # Convertir version a numero para validar
        $versionNum = [double]$version
        if ($versionNum -ge 12) { # Excel 2007 o superior
            return $true, $version
        } else {
            return $false, $version
        }
    }
    catch {
        return $false, $null
    }
}

# Funcion para ejecutar proceso con timeout
function Invoke-ProcessWithTimeout {
    param(
        [string]$FilePath,
        [string]$Arguments,
        [int]$TimeoutSeconds = 30
    )
    
    try {
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $FilePath
        $processInfo.Arguments = $Arguments
        $processInfo.RedirectStandardError = $true
        $processInfo.RedirectStandardOutput = $true
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        
        # Iniciar el proceso
        $process.Start() | Out-Null
        
        # Esperar con timeout
        if ($process.WaitForExit($TimeoutSeconds * 1000)) {
            # Proceso completado dentro del timeout
            $stdout = $process.StandardOutput.ReadToEnd()
            $stderr = $process.StandardError.ReadToEnd()
            
            return @{
                Success = $true
                ExitCode = $process.ExitCode
                Stdout = $stdout
                Stderr = $stderr
                TimedOut = $false
            }
        } else {
            # Timeout - matar el proceso
            Write-Host "  [TIMEOUT] El proceso excedio el tiempo de espera ($TimeoutSeconds segundos)" -ForegroundColor Red
            try {
                $process.Kill()
                Start-Sleep -Milliseconds 500
            } catch {
                # Intentar forzar la terminacion
                try {
                    Stop-Process -Id $process.Id -Force -ErrorAction SilentlyContinue
                } catch {}
            }
            
            return @{
                Success = $false
                ExitCode = -1
                Stdout = ""
                Stderr = "TIMEOUT: El proceso excedio el tiempo de espera de $TimeoutSeconds segundos."
                TimedOut = $true
            }
        }
    }
    catch {
        return @{
            Success = $false
            ExitCode = -1
            Stdout = ""
            Stderr = "ERROR: $($_.Exception.Message)"
            TimedOut = $false
        }
    }
    finally {
        # Asegurarse de que el proceso este cerrado
        if ($process -and !$process.HasExited) {
            try {
                $process.Kill()
            } catch {}
        }
    }
}

# Funcion para verificar conexion a red
function Test-NetworkConnection {
    param(
        [string]$Hosting,
        [int]$Port
    )
    
    try {
        Write-Host "Verificando conectividad de red..." -ForegroundColor Yellow -NoNewline
        
        # Intentar conexion TCP
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $connectResult = $tcpClient.BeginConnect($Hosting, $Port, $null, $null)
        $waitResult = $connectResult.AsyncWaitHandle.WaitOne(10000, $false) # 10 segundos timeout
        
        if ($waitResult) {
            $tcpClient.EndConnect($connectResult)
            $tcpClient.Close()
            Write-Host " [OK]" -ForegroundColor Green
            return $true
        } else {
            $tcpClient.Close()
            Write-Host " [FALLO]" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host " [FALLO]" -ForegroundColor Red
        return $false
    }
}

# Funcion para filtrar warnings de Java
function Remove-JavaWarnings {
    param(
        [string]$Text
    )
    
    if ([string]::IsNullOrEmpty($Text)) {
        return $Text
    }
    
    # Filtrar warnings especificos de Java 17+
    $patterns = @(
        'WARNING: A restricted method in java\.lang\.System has been called',
        'WARNING: java\.lang\.System::[a-zA-Z]+ has been called',
        'WARNING: Please consider reporting this to the maintainers',
        'WARNING: Use --illegal-access=warn to enable warnings',
        'WARNING: All illegal access operations will be denied',
        'WARNING: An illegal reflective access operation has occurred',
        'WARNING: Illegal reflective access by',
        'WARNING: Using incubator modules'
    )
    
    $result = $Text
    foreach ($pattern in $patterns) {
        # Usar regex para eliminar lineas completas que contengan estos warnings
        $result = [regex]::Replace($result, "(?m)^\s*$pattern.*$[\r\n]*", "")
    }
    
    # Eliminar lineas vacias multiples
    $result = [regex]::Replace($result, "(?m)^\s*$[\r\n]+", "`r`n")
    
    return $result.Trim()
}

try {
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " Ejecutor de Consultas SQL - Oracle" -ForegroundColor Cyan
    Write-Host " Oracle SQLcl con exportacion a CSV/Excel" -ForegroundColor Cyan
    Write-Host " Usando Microsoft Excel para conversion" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    # ========================================
    # Solicitar credenciales y datos de conexion
    # ========================================
    
    # Usuario
    do {
        $usuario = Read-Host "Ingrese el usuario de Oracle"
        if ([string]::IsNullOrWhiteSpace($usuario)) {
            Write-Host "[ERROR] El usuario no puede estar vacio" -ForegroundColor Red
            Write-Host ""
        }
    } while ([string]::IsNullOrWhiteSpace($usuario))

    # Contrasena (enmascarada)
    do {
        $securePassword = Read-Host "Ingrese la contrasena" -AsSecureString
        # Convertir SecureString a texto plano para usar con SQLcl
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
        $password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        
        if ([string]::IsNullOrWhiteSpace($password)) {
            Write-Host "[ERROR] La contrasena no puede estar vacia" -ForegroundColor Red
            Write-Host ""
        }
    } while ([string]::IsNullOrWhiteSpace($password))

    # Host
    do {
        $host_db = Read-Host "Ingrese el host (ej: localhost)"
        if ([string]::IsNullOrWhiteSpace($host_db)) {
            Write-Host "[ERROR] El host no puede estar vacio" -ForegroundColor Red
            Write-Host ""
        }
    } while ([string]::IsNullOrWhiteSpace($host_db))

    # Puerto
    do {
        $puerto = Read-Host "Ingrese el puerto (ej: 1521)"
        if ([string]::IsNullOrWhiteSpace($puerto)) {
            Write-Host "[ERROR] El puerto no puede estar vacio" -ForegroundColor Red
            Write-Host ""
            continue
        }
        
        # Validar que sea numerico
        if (-not ($puerto -match '^\d+$')) {
            Write-Host "[ERROR] El puerto debe ser un numero valido" -ForegroundColor Red
            Write-Host ""
            $puerto = $null
            continue
        }
        
        # Convertir a entero y validar rango
        $puertoInt = [int]$puerto
        if ($puertoInt -lt 1 -or $puertoInt -gt 65535) {
            Write-Host "[ERROR] El puerto debe estar entre 1 y 65535" -ForegroundColor Red
            Write-Host ""
            $puerto = $null
        }
    } while ([string]::IsNullOrWhiteSpace($puerto))

    # SID o Service Name
    do {
        $sidService = Read-Host "Ingrese el SID o Service Name"
        if ([string]::IsNullOrWhiteSpace($sidService)) {
            Write-Host "[ERROR] El SID o Service Name no puede estar vacio" -ForegroundColor Red
            Write-Host ""
        }
    } while ([string]::IsNullOrWhiteSpace($sidService))

    # Formato de salida
    do {
        $tipoSalida = Read-Host "Formato de salida (1=CSV, 2=XLSX/Excel) [Por defecto: 1]"
        if ([string]::IsNullOrWhiteSpace($tipoSalida)) {
            $tipoSalida = "1"
        }
        
        if ($tipoSalida -eq "1") {
            $extension = "csv"
            $formato = "csv"
            $formatoDisplay = "CSV"
            $exportarExcel = $false
        }
        elseif ($tipoSalida -eq "2") {
            $extension = "xlsx"
            $formato = "xlsx"
            $formatoDisplay = "XLSX (Excel)"
            $exportarExcel = $true
            
            # Verificar si Excel esta instalado
            $excelInfo = Test-ExcelInstalled
            if (-not $excelInfo[0]) {
                Write-Host "[ERROR] Microsoft Excel no esta instalado o es anterior a 2007" -ForegroundColor Red
                if ($excelInfo[1]) {
                    Write-Host "Version encontrada: $($excelInfo[1]) (se requiere Excel 2007 o superior)" -ForegroundColor Yellow
                } else {
                    Write-Host "Microsoft Excel no fue encontrado en el sistema" -ForegroundColor Yellow
                }
                Write-Host ""
                Write-Host "Por favor, instale Microsoft Excel 2007 o superior para usar esta funcion." -ForegroundColor Yellow
                Write-Host "Puede seleccionar la opcion 1 para exportar a CSV en lugar de Excel." -ForegroundColor Yellow
                throw "Microsoft Excel no disponible"
            }
            
            Write-Host "[OK] Microsoft Excel $($excelInfo[1]) detectado" -ForegroundColor Green
        }
        else {
            Write-Host "[ERROR] Opcion invalida. Debe ser 1 o 2" -ForegroundColor Red
            Write-Host ""
            $tipoSalida = $null
        }
    } while ([string]::IsNullOrWhiteSpace($tipoSalida))

    Write-Host ""
    Write-Host "Configuracion recibida:" -ForegroundColor Yellow
    Write-Host "  Usuario: $usuario"
    Write-Host "  Host: $host_db"
    Write-Host "  Puerto: $puerto"
    Write-Host "  SID/Service: $sidService"
    Write-Host "  Formato: $formatoDisplay"
    Write-Host ""

    # ========================================
    # Verificar conectividad de red primero
    # ========================================
    Write-Host "Validando conectividad de red hacia ${host_db}:${puerto}..." -ForegroundColor Yellow
    $networkTest = Test-NetworkConnection -Hosting $host_db -Port $puertoInt
    
    if (-not $networkTest) {
        Write-Host "[ERROR] No se puede conectar a $host_db en el puerto $puerto" -ForegroundColor Red
        Write-Host ""
        Write-Host "Por favor verifique:"
        Write-Host "  1. Que el servidor Oracle este en ejecucion"
        Write-Host "  2. Que el host y puerto sean correctos"
        Write-Host "  3. Que no haya firewall bloqueando la conexion"
        Write-Host "  4. Que tenga acceso de red al servidor"
        Write-Host ""
        throw "Error de conectividad de red"
    }

    # ========================================
    # Verificar estructura de carpetas
    # ========================================
    
    $dirActual = Split-Path -Parent $MyInvocation.MyCommand.Path
    $dirConsultas = Join-Path $dirActual "consultas"
    $dirResultados = Join-Path $dirActual "resultados"

    if (-not (Test-Path $dirConsultas)) {
        Write-Host "[ERROR] No se encontro la carpeta 'consultas'" -ForegroundColor Red
        Write-Host ""
        Write-Host "Por favor, cree la carpeta 'consultas' en el mismo directorio donde esta este script"
        Write-Host "y coloque sus archivos .sql dentro de ella."
        Write-Host ""
        throw "Carpeta 'consultas' no encontrada"
    }

    if (-not (Test-Path $dirResultados)) {
        Write-Host "[ERROR] No se encontro la carpeta 'resultados'" -ForegroundColor Red
        Write-Host ""
        Write-Host "Por favor, cree la carpeta 'resultados' en el mismo directorio donde esta este script."
        Write-Host "Esta carpeta se utilizara para guardar los resultados de las consultas."
        Write-Host ""
        throw "Carpeta 'resultados' no encontrada"
    }

    # ========================================
    # Buscar instalacion de Oracle SQLcl
    # ========================================
    
    $sqlclPaths = @(
        "C:\oracle\sqlcl\bin\sql.exe",
        "C:\Program Files\Oracle\sqlcl\bin\sql.exe",
        "$env:USERPROFILE\sqlcl\bin\sql.exe",
        "$env:ORACLE_HOME\sqlcl\bin\sql.exe"
    )

    $sqlclPath = $null
    foreach ($path in $sqlclPaths) {
        if (Test-Path $path) {
            $sqlclPath = $path
            break
        }
    }

    # Si no se encuentra en rutas conocidas, buscar en PATH
    if (-not $sqlclPath) {
        try {
            $sqlInPath = Get-Command sql.exe -ErrorAction Stop
            $sqlclPath = $sqlInPath.Source
        }
        catch {
            # No esta en PATH
        }
    }

    if (-not $sqlclPath) {
        Write-Host "[ERROR] No se pudo encontrar Oracle SQLcl instalado" -ForegroundColor Red
        Write-Host ""
        Write-Host "Rutas buscadas:"
        foreach ($path in $sqlclPaths) {
            Write-Host "  - $path"
        }
        Write-Host "  - Variable PATH del sistema"
        Write-Host ""
        Write-Host "Por favor, descargue e instale Oracle SQLcl desde:"
        Write-Host "https://www.oracle.com/database/sqldeveloper/technologies/sqlcl/download/" -ForegroundColor Cyan
        Write-Host ""
        throw "Oracle SQLcl no encontrado"
    }

    Write-Host "Oracle SQLcl encontrado en: $sqlclPath" -ForegroundColor Green
    Write-Host ""

    # ========================================
    # SOLUCION: Configurar variables de entorno para Java
    # ========================================
    Write-Host "Configurando variables de entorno para evitar warnings de Java..." -ForegroundColor Yellow
    
    # Crear archivo batch temporal para ejecutar SQLcl con las opciones correctas
    $sqlclBatPath = Join-Path $env:TEMP "sqlcl_wrapper_$(Get-Random).bat"
    
    # Obtener el directorio de SQLcl
    $sqlclDir = Split-Path $sqlclPath -Parent
    
    # Crear el batch que ejecutara SQLcl con las opciones de Java correctas
    @"
@echo off
setlocal

REM Configurar variables de entorno para evitar warnings de Java 17+
set JAVA_TOOL_OPTIONS=-Duser.language=en -Duser.country=US
set JDK_JAVA_OPTIONS=--add-opens=java.base/java.lang=ALL-UNNAMED --add-opens=java.base/java.io=ALL-UNNAMED --add-opens=java.base/java.util=ALL-UNNAMED

REM Cambiar al directorio de SQLcl
cd /d "$sqlclDir"

REM Ejecutar SQLcl con los argumentos pasados
call sql.exe %*
"@ | Out-File -FilePath $sqlclBatPath -Encoding ASCII

    Write-Host "[OK] Wrapper creado para evitar warnings de Java" -ForegroundColor Green
    Write-Host ""

    # ========================================
    # Validar conexion a Oracle (CON TIMEOUT)
    # ========================================
    
    $connectionString = "${usuario}/${password}@${host_db}:${puerto}/${sidService}"
    
    Write-Host "Validando conexion a Oracle (timeout: 30 segundos)..." -ForegroundColor Yellow
    Write-Host ""

    # Crear script temporal para probar conexion
    $testScript = Join-Path $env:TEMP "test_connection_$(Get-Random).sql"
    @"
SELECT 'CONNECTION_OK' FROM DUAL;
"@ | Out-File -FilePath $testScript -Encoding UTF8

    # Intentar conexion con timeout usando el wrapper
    $testResult = Invoke-ProcessWithTimeout -FilePath $sqlclBatPath -Arguments "-S $connectionString @`"$testScript`"" -TimeoutSeconds 30

    # Limpiar archivo temporal
    Remove-Item $testScript -ErrorAction SilentlyContinue

    # Filtrar warnings de Java del resultado
    $filteredStderr = Remove-JavaWarnings -Text $testResult.Stderr
    
    if (-not $testResult.Success -or $testResult.TimedOut) {
        Write-Host "[ERROR] No se pudo establecer conexion con Oracle" -ForegroundColor Red
        Write-Host ""
        
        if ($testResult.TimedOut) {
            Write-Host "La conexion se agoto (timeout de 30 segundos)" -ForegroundColor Yellow
        } elseif ($filteredStderr -and $filteredStderr.Trim() -ne "") {
            Write-Host "Error detallado:" -ForegroundColor Yellow
            Write-Host $filteredStderr -ForegroundColor Red
        }
        
        Write-Host ""
        Write-Host "Verifique los siguientes datos:"
        Write-Host "  - Usuario: $usuario"
        Write-Host "  - Host: $host_db"
        Write-Host "  - Puerto: $puerto"
        Write-Host "  - SID/Service: $sidService"
        Write-Host ""
        Write-Host "Posibles causas:"
        Write-Host "  - Credenciales incorrectas"
        Write-Host "  - Servidor Oracle no accesible"
        Write-Host "  - Firewall bloqueando la conexion"
        Write-Host "  - SID o Service Name incorrecto"
        Write-Host "  - El servicio Oracle no esta en ejecucion"
        Write-Host ""
        
        # Limpiar archivo batch temporal
        Remove-Item $sqlclBatPath -ErrorAction SilentlyContinue
        
        throw "Error de conexion a Oracle"
    }

    # Verificar la respuesta ignorando warnings
    if ($testResult.Stdout -notlike "*CONNECTION_OK*") {
        # Si hay warnings pero la conexion fue exitosa, verificar en stderr filtrado
        if ($filteredStderr -and $filteredStderr.Trim() -ne "") {
            Write-Host "[ADVERTENCIA] Hubo advertencias durante la conexion:" -ForegroundColor Yellow
            Write-Host $filteredStderr -ForegroundColor DarkYellow
            Write-Host ""
        }
        
        # Verificar si realmente fallo o solo hay warnings
        Write-Host "[ERROR] Respuesta inesperada del servidor Oracle" -ForegroundColor Red
        Write-Host ""
        Write-Host "Respuesta recibida:" -ForegroundColor Yellow
        Write-Host $testResult.Stdout -ForegroundColor Red
        
        # Limpiar archivo batch temporal
        Remove-Item $sqlclBatPath -ErrorAction SilentlyContinue
        
        throw "Error en validacion de conexion"
    }

    Write-Host "[OK] Conexion establecida correctamente" -ForegroundColor Green
    
    # Mostrar warnings si los hay (filtrados)
    if ($filteredStderr -and $filteredStderr.Trim() -ne "") {
        Write-Host "[INFO] Nota: Se ignoraron warnings de Java durante la conexion" -ForegroundColor Gray
    }
    
    Write-Host ""

    # ========================================
    # Buscar archivos SQL
    # ========================================
    
    $archivosSql = Get-ChildItem -Path $dirConsultas -Filter "*.sql"
    $count = $archivosSql.Count

    if ($count -eq 0) {
        Write-Host "[ADVERTENCIA] No se encontraron archivos .sql en la carpeta 'consultas'" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Por favor, agregue sus consultas SQL en la carpeta 'consultas' y vuelva a ejecutar el script."
        Write-Host ""
        
        # Limpiar archivo batch temporal
        Remove-Item $sqlclBatPath -ErrorAction SilentlyContinue
        
        return
    }

    Write-Host "Se encontraron $count archivo(s) SQL para procesar" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Iniciando procesamiento..." -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    # ========================================
    # Procesar cada archivo SQL (CON TIMEOUT)
    # ========================================
    
    $procesados = 0
    $errores = 0
    $timeouts = 0

    foreach ($archivo in $archivosSql) {
        $nombreArchivo = $archivo.Name
        $nombreBase = $archivo.BaseName
        
        # Generar timestamp
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        
        $archivoSalida = "${nombreBase}_${timestamp}.${extension}"
        $rutaSalida = Join-Path $dirResultados $archivoSalida
        
        Write-Host "Procesando: $nombreArchivo" -ForegroundColor White
        Write-Host "  > Salida: $archivoSalida" -ForegroundColor Gray
        
        # Crear archivo temporal para el CSV (siempre se exporta primero a CSV)
        $archivoTempCsv = Join-Path $env:TEMP "${nombreBase}_${timestamp}_temp.csv"
        
        # Crear script temporal con comandos SQLcl para exportar a CSV
        $wrapperScript = Join-Path $env:TEMP "wrapper_$(Get-Random).sql"
        
        @"
SET ECHO OFF
SET FEEDBACK OFF
SET PAGESIZE 0
SET LINESIZE 32767
SET TRIMSPOOL ON
SET SQLFORMAT csv
SET TERMOUT OFF
SPOOL $archivoTempCsv
@"$($archivo.FullName)"
SPOOL OFF
EXIT;
"@ | Out-File -FilePath $wrapperScript -Encoding UTF8
        
        # Ejecutar consulta CON TIMEOUT (30 minutos = 1800 segundos) usando el wrapper
        Write-Host "  > Ejecutando consulta (timeout: 30 minutos)..." -ForegroundColor Gray
        $queryResult = Invoke-ProcessWithTimeout -FilePath $sqlclBatPath -Arguments "-S $connectionString @`"$wrapperScript`"" -TimeoutSeconds 1800

        # Limpiar script temporal
        Remove-Item $wrapperScript -ErrorAction SilentlyContinue

        # Filtrar warnings de Java
        $filteredQueryStderr = Remove-JavaWarnings -Text $queryResult.Stderr
        
        # Verificar si hubo timeout
        if ($queryResult.TimedOut) {
            Write-Host "  [TIMEOUT] La consulta excedio el tiempo maximo de ejecucion (30 minutos)" -ForegroundColor Red
            $timeouts++
            $errores++
            
            # Intentar limpiar archivos temporales
            if (Test-Path $archivoTempCsv) {
                Remove-Item $archivoTempCsv -ErrorAction SilentlyContinue
            }
            Write-Host ""
            continue
        }

        # Verificar si la consulta se ejecuto correctamente (ignorando warnings)
        if (-not (Test-Path $archivoTempCsv)) {
            Write-Host "  [ERROR] Fallo la ejecucion de la consulta - No se generó el archivo CSV" -ForegroundColor Red
            
            # Mostrar solo errores reales (filtrados de warnings)
            if ($filteredQueryStderr -and $filteredQueryStderr.Trim() -ne "") {
                $errorMsg = $filteredQueryStderr.Trim()
                if ($errorMsg.Length -gt 300) {
                    $errorMsg = $errorMsg.Substring(0, 300) + "..."
                }
                Write-Host "  Detalles: $errorMsg" -ForegroundColor DarkRed
            }
            
            $errores++
            Write-Host ""
            continue
        }

        # Verificar que el CSV no este vacio o solo contenga encabezados
        $csvContent = Get-Content $archivoTempCsv -Raw
        $lineas = $csvContent -split "`n" | Where-Object { $_.Trim() -ne "" }
        
        if ($lineas.Count -le 1) {
            Write-Host "  [ADVERTENCIA] La consulta no devolvio datos o solo contiene encabezados" -ForegroundColor Yellow
            
            # Mostrar warnings si los hay
            if ($filteredQueryStderr -and $filteredQueryStderr.Trim() -ne "") {
                Write-Host "  Nota: $($filteredQueryStderr.Trim())" -ForegroundColor DarkYellow
            }
            
            Remove-Item $archivoTempCsv -ErrorAction SilentlyContinue
            $procesados++
            Write-Host ""
            continue
        }

        try {
            if ($exportarExcel) {
                # Exportar a Excel (XLSX) usando Microsoft Excel
                Write-Host "  > Convirtiendo CSV a Excel usando Microsoft Office..." -ForegroundColor Gray
                
                # Crear objeto Excel
                $excel = New-Object -ComObject "Excel.Application"
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                $excel.ScreenUpdating = $false
                $excel.AskToUpdateLinks = $false
                
                try {
                    # Abrir el archivo CSV
                    Write-Host "  > Abriendo CSV en Excel..." -ForegroundColor Gray
                    $workbook = $excel.Workbooks.Open($archivoTempCsv)
                    
                    # Aplicar formato basico
                    $worksheet = $workbook.Worksheets.Item(1)
                    
                    # Autoajustar columnas
                    $usedRange = $worksheet.UsedRange
                    $usedRange.EntireColumn.AutoFit() | Out-Null
                    
                    # Formato de tabla (solo Excel 2007+)
                    $listObject = $worksheet.ListObjects.Add(
                        [Microsoft.Office.Interop.Excel.XlListObjectSourceType]::xlSrcRange,
                        $usedRange,
                        $null,
                        [Microsoft.Office.Interop.Excel.XlYesNoGuess]::xlYes
                    )
                    
                    # Aplicar estilo de tabla (estilo 6 es uno basico)
                    $listObject.TableStyle = "TableStyleMedium2"
                    
                    # Congelar paneles (primera fila)
                    $worksheet.Activate()
                    $excel.ActiveWindow.SplitRow = 1
                    $excel.ActiveWindow.FreezePanes = $true
                    
                    # Guardar como XLSX
                    Write-Host "  > Guardando como XLSX..." -ForegroundColor Gray
                    
                    # Formato XLSX (51 = xlOpenXMLWorkbook - .xlsx)
                    $xlFileFormat = 51
                    
                    $workbook.SaveAs($rutaSalida, $xlFileFormat)
                    
                    Write-Host "  [OK] Archivo Excel generado: $archivoSalida" -ForegroundColor Green
                    
                    # Mostrar warnings si los hubo durante la consulta
                    if ($filteredQueryStderr -and $filteredQueryStderr.Trim() -ne "") {
                        Write-Host "  [INFO] Nota: Se ignoraron warnings de Java durante la ejecucion" -ForegroundColor Gray
                    }
                }
                catch {
                    Write-Host "  [ERROR] Error al convertir a Excel: $($_.Exception.Message)" -ForegroundColor Red
                    throw
                }
                finally {
                    # Cerrar todo correctamente
                    if ($workbook) {
                        try { 
                            $workbook.Close($false) 
                        } catch {}
                        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
                    }
                    
                    $excel.Quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                    
                    # Matar procesos de Excel residuales
                    Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                }
            }
            else {
                # Solo CSV - mover el archivo temporal a la ubicacion final
                Move-Item -Path $archivoTempCsv -Destination $rutaSalida -Force
                Write-Host "  [OK] Archivo CSV generado: $archivoSalida" -ForegroundColor Green
                
                # Mostrar warnings si los hubo durante la consulta
                if ($filteredQueryStderr -and $filteredQueryStderr.Trim() -ne "") {
                    Write-Host "  [INFO] Nota: Se ignoraron warnings de Java durante la ejecucion" -ForegroundColor Gray
                }
            }
            
            $procesados++
        }
        catch {
            Write-Host "  [ERROR] Error al procesar los datos: $($_.Exception.Message)" -ForegroundColor Red
            
            # Si fallo la conversion a Excel, guardar como CSV como respaldo
            if ($exportarExcel -and (Test-Path $archivoTempCsv)) {
                $backupCsv = Join-Path $dirResultados "${nombreBase}_${timestamp}_backup.csv"
                Move-Item -Path $archivoTempCsv -Destination $backupCsv -Force
                Write-Host "  [INFO] Datos guardados como CSV de respaldo: $(Split-Path $backupCsv -Leaf)" -ForegroundColor Yellow
            }
            
            $errores++
        }
        finally {
            # Limpiar archivo temporal si existe
            if (Test-Path $archivoTempCsv) {
                Remove-Item $archivoTempCsv -ErrorAction SilentlyContinue
            }
        }
        
        Write-Host ""
    }

    # ========================================
    # Limpiar archivo batch temporal
    # ========================================
    Remove-Item $sqlclBatPath -ErrorAction SilentlyContinue

    # ========================================
    # Resumen final
    # ========================================
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Procesamiento completado" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Archivos procesados: $procesados" -ForegroundColor Green
    Write-Host "Errores encontrados: $errores" -ForegroundColor $(if ($errores -gt 0) { "Red" } else { "Green" })
    if ($timeouts -gt 0) {
        Write-Host "Timeouts de consulta: $timeouts" -ForegroundColor Yellow
    }
    Write-Host ""

    if ($errores -gt 0) {
        Write-Host "[NOTA] Algunos archivos no se procesaron correctamente." -ForegroundColor Yellow
        Write-Host "Verifique:"
        Write-Host "  - Las credenciales de conexion"
        Write-Host "  - La sintaxis de las consultas SQL"
        Write-Host "  - La conectividad con el servidor Oracle"
        Write-Host "  - Que Excel no este bloqueado por otro proceso"
        Write-Host "  - Si hay timeouts, considere optimizar las consultas largas"
        Write-Host ""
    }

    Write-Host "Resultados guardados en: $dirResultados" -ForegroundColor Cyan
    
    if ($exportarExcel -and $procesados -gt 0) {
        Write-Host ""
        Write-Host "[INFO] Archivos Excel generados con las siguientes caracteristicas:" -ForegroundColor Cyan
        Write-Host "  - Tabla formateada con estilo profesional" -ForegroundColor Gray
        Write-Host "  - Encabezados congelados para facil navegacion" -ForegroundColor Gray
        Write-Host "  - Columnas autoajustadas al contenido" -ForegroundColor Gray
        Write-Host "  - Formato XLSX compatible con Excel 2007+" -ForegroundColor Gray
    }
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Script finalizado exitosamente" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
}
catch {
    # ========================================
    # BLOQUE CATCH - Manejo de errores
    # ========================================
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "[ERROR CRITICO] Se produjo un error" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
    Write-Host "Detalles del error:" -ForegroundColor Yellow
    Write-Host $_.Exception.Message -ForegroundColor Red
    if ($_.ScriptStackTrace) {
        Write-Host "Trace: $($_.ScriptStackTrace)" -ForegroundColor DarkGray
    }
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "El script finalizo con errores" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host ""
}
finally {
    # ========================================
    # BLOQUE FINALLY - Siempre se ejecuta
    # ========================================
    # Limpiar procesos SQLcl que puedan haber quedado
    Get-Process | Where-Object { $_.ProcessName -eq "sql" } | Stop-Process -Force -ErrorAction SilentlyContinue
    
    # Limpiar archivo batch temporal si existe
    $batFiles = Get-ChildItem "$env:TEMP\sqlcl_wrapper_*.bat" -ErrorAction SilentlyContinue
    foreach ($batFile in $batFiles) {
        Remove-Item $batFile.FullName -ErrorAction SilentlyContinue
    }
    
    Write-Host "Presione ENTER para salir..." -ForegroundColor Cyan
    $null = Read-Host
}
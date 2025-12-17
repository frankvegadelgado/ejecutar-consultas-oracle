# ========================================
# Ejecutor de Consultas SQL - Oracle
# Oracle SQLcl
# PowerShell Script
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

try {
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " Ejecutor de Consultas SQL - Oracle" -ForegroundColor Cyan
    Write-Host " Oracle SQLcl" -ForegroundColor Cyan
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
        $tipoSalida = Read-Host "Formato de salida (1=CSV, 2=JSON) [Por defecto: 1]"
        if ([string]::IsNullOrWhiteSpace($tipoSalida)) {
            $tipoSalida = "1"
        }
        
        if ($tipoSalida -eq "1") {
            $extension = "csv"
            $formato = "csv"
        }
        elseif ($tipoSalida -eq "2") {
            $extension = "json"
            $formato = "json"
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
    Write-Host "  Formato: $($formato.ToUpper())"
    Write-Host ""

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
    # Validar conexion a Oracle
    # ========================================
    
    $connectionString = "${usuario}/${password}@${host_db}:${puerto}/${sidService}"
    
    Write-Host "Validando conexion a Oracle..." -ForegroundColor Yellow
    Write-Host ""

    # Crear script temporal para probar conexion
    $testScript = Join-Path $env:TEMP "test_connection_$(Get-Random).sql"
    @"
SET ECHO OFF
SET FEEDBACK OFF
SET HEADING OFF
SELECT 'CONNECTION_OK' FROM DUAL;
EXIT;
"@ | Out-File -FilePath $testScript -Encoding UTF8

    # Intentar conexion
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = $sqlclPath
    $processInfo.Arguments = "-S $connectionString @`"$testScript`""
    $processInfo.RedirectStandardError = $true
    $processInfo.RedirectStandardOutput = $true
    $processInfo.UseShellExecute = $false
    $processInfo.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processInfo
    $process.Start() | Out-Null
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    # Limpiar archivo temporal
    Remove-Item $testScript -ErrorAction SilentlyContinue

    if ($process.ExitCode -ne 0 -or $stdout -notlike "*CONNECTION_OK*") {
        Write-Host "[ERROR] No se pudo establecer conexion con Oracle" -ForegroundColor Red
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
        Write-Host ""
        if ($stderr) {
            Write-Host "Detalles del error:" -ForegroundColor Yellow
            Write-Host $stderr -ForegroundColor Red
            Write-Host ""
        }
        throw "Error de conexion a Oracle"
    }

    Write-Host "[OK] Conexion establecida correctamente" -ForegroundColor Green
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
        return
    }

    Write-Host "Se encontraron $count archivo(s) SQL para procesar" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Iniciando procesamiento..." -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""

    # ========================================
    # Procesar cada archivo SQL
    # ========================================
    
    $procesados = 0
    $errores = 0

    foreach ($archivo in $archivosSql) {
        $nombreArchivo = $archivo.Name
        $nombreBase = $archivo.BaseName
        
        # Generar timestamp
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        
        $archivoSalida = "${nombreBase}_${timestamp}.${extension}"
        $rutaSalida = Join-Path $dirResultados $archivoSalida
        
        Write-Host "Procesando: $nombreArchivo" -ForegroundColor White
        Write-Host "  > Salida: $archivoSalida" -ForegroundColor Gray
        
        # Crear script temporal con comandos SQLcl
        $wrapperScript = Join-Path $env:TEMP "wrapper_$(Get-Random).sql"
        
        if ($formato -eq "csv") {
            @"
SET ECHO OFF
SET FEEDBACK OFF
SET PAGESIZE 0
SET LINESIZE 32767
SET TRIMSPOOL ON
SET SQLFORMAT csv
SPOOL $rutaSalida
@"$($archivo.FullName)"
SPOOL OFF
EXIT;
"@ | Out-File -FilePath $wrapperScript -Encoding UTF8
        }
        else {
            # JSON format
            @"
SET ECHO OFF
SET FEEDBACK OFF
SET PAGESIZE 0
SET SQLFORMAT json
SPOOL $rutaSalida
@"$($archivo.FullName)"
SPOOL OFF
EXIT;
"@ | Out-File -FilePath $wrapperScript -Encoding UTF8
        }
        
        # Ejecutar consulta
        $queryProcessInfo = New-Object System.Diagnostics.ProcessStartInfo
        $queryProcessInfo.FileName = $sqlclPath
        $queryProcessInfo.Arguments = "-S $connectionString @`"$wrapperScript`""
        $queryProcessInfo.RedirectStandardError = $true
        $queryProcessInfo.RedirectStandardOutput = $true
        $queryProcessInfo.UseShellExecute = $false
        $queryProcessInfo.CreateNoWindow = $true

        $queryProcess = New-Object System.Diagnostics.Process
        $queryProcess.StartInfo = $queryProcessInfo
        $queryProcess.Start() | Out-Null
        $queryStdout = $queryProcess.StandardOutput.ReadToEnd()
        $queryStderr = $queryProcess.StandardError.ReadToEnd()
        $queryProcess.WaitForExit()

        # Limpiar script temporal
        Remove-Item $wrapperScript -ErrorAction SilentlyContinue

        if ($queryProcess.ExitCode -eq 0 -and (Test-Path $rutaSalida)) {
            Write-Host "  [OK] Consulta ejecutada correctamente" -ForegroundColor Green
            $procesados++
        }
        else {
            Write-Host "  [ERROR] Fallo la ejecucion de la consulta" -ForegroundColor Red
            if ($queryStderr) {
                Write-Host "  Detalles: $($queryStderr.Substring(0, [Math]::Min(100, $queryStderr.Length)))" -ForegroundColor DarkRed
            }
            $errores++
        }
        Write-Host ""
    }

    # ========================================
    # Resumen final
    # ========================================
    
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Procesamiento completado" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Archivos procesados: $procesados" -ForegroundColor Green
    Write-Host "Errores encontrados: $errores" -ForegroundColor $(if ($errores -gt 0) { "Red" } else { "Green" })
    Write-Host ""

    if ($errores -gt 0) {
        Write-Host "[NOTA] Algunos archivos no se procesaron correctamente." -ForegroundColor Yellow
        Write-Host "Verifique:"
        Write-Host "  - Las credenciales de conexion"
        Write-Host "  - La sintaxis de las consultas SQL"
        Write-Host "  - La conectividad con el servidor Oracle"
        Write-Host ""
    }

    Write-Host "Resultados guardados en: $dirResultados" -ForegroundColor Cyan
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
    Write-Host "Presione ENTER para salir..." -ForegroundColor Cyan
    $null = Read-Host
}
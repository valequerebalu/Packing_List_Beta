@echo off
echo ======================================
echo  Diagnostico de conexion Beta Packing
echo ======================================
echo.

echo --- 1. Prueba DNS ---
nslookup xythcnwjtbfwwcxxprbp.supabase.co
echo.

echo --- 2. Prueba Ping ---
ping -n 3 xythcnwjtbfwwcxxprbp.supabase.co
echo.

echo --- 3. Prueba Puerto 443 (HTTPS) ---
echo Intentando conectar al puerto 443...
powershell -Command "try { $tcp = New-Object System.Net.Sockets.TcpClient; $tcp.Connect('xythcnwjtbfwwcxxprbp.supabase.co', 443); Write-Host '[OK] Puerto 443 accesible'; $tcp.Close() } catch { Write-Host '[FAIL] Puerto 443 bloqueado:' $_.Exception.Message }"
echo.

echo --- 4. Prueba HTTP a Supabase ---
powershell -Command "try { $r = Invoke-WebRequest -Uri 'https://xythcnwjtbfwwcxxprbp.supabase.co/rest/v1/' -UseBasicParsing -TimeoutSec 10; Write-Host '[OK] HTTP Status:' $r.StatusCode } catch { Write-Host '[FAIL] HTTP fallo:' $_.Exception.Message }"
echo.

echo --- 5. Verificar archivo .env ---
if exist "%~dp0.env" (
    echo [OK] .env encontrado en directorio del exe
    type "%~dp0.env"
) else (
    echo [WARN] .env NO encontrado en directorio del exe
    echo Buscando en carpeta temporal...
    if exist "%TEMP%\_MEI*\\.env" (
        echo [OK] .env encontrado en MEIPASS
    ) else (
        echo [FAIL] .env no encontrado en ninguna ubicacion
    )
)

echo.
echo ======================================
echo  Fin del diagnostico
echo ======================================
pause

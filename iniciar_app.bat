@echo off
cd /d "%~dp0"
title Generador de manuales OyM

rem === Si el .bat no encuentra Python solo, pega aqui la ruta completa de python.exe ===
rem Ejemplo: set "PY_MANUAL=C:\Users\TuUsuario\AppData\Local\Programs\Python\Python312\python.exe"
set "PY_MANUAL="

rem --- Primera vez: crea el entorno e instala todo (necesita internet) ---
if exist ".venv\Scripts\python.exe" goto :arrancar

set "PY="
set "PYARGS="
if defined PY_MANUAL set "PY=%PY_MANUAL%"
if defined PY goto :crear

py -3 --version >nul 2>&1
if not errorlevel 1 set "PY=py"
if not errorlevel 1 set "PYARGS=-3"
if defined PY goto :crear

python --version >nul 2>&1
if not errorlevel 1 set "PY=python"
if defined PY goto :crear

for /d %%D in ("%LocalAppData%\Programs\Python\Python3*") do if exist "%%D\python.exe" set "PY=%%D\python.exe"
if defined PY goto :crear

echo.
echo No se encontro Python en este equipo.
echo Abre una terminal donde si te funcione la app y escribe:  where python
echo Luego pega esa ruta en la linea PY_MANUAL de este archivo.
echo.
pause
exit /b 1

:crear
echo Preparando la aplicacion por primera vez. Puede tardar unos minutos...
"%PY%" %PYARGS% -m venv .venv
if errorlevel 1 goto :fallo
".venv\Scripts\python.exe" -m pip install --upgrade pip -q
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :fallo
goto :arrancar

:fallo
echo.
echo Fallo la preparacion. Revisa la conexion a internet y vuelve a intentar.
if exist ".venv" rmdir /s /q ".venv"
pause
exit /b 1

:arrancar
rem --- Evita que Streamlit pida un correo la primera vez ---
if exist "%USERPROFILE%\.streamlit\credentials.toml" goto :cred_ok
mkdir "%USERPROFILE%\.streamlit" 2>nul
echo [general]> "%USERPROFILE%\.streamlit\credentials.toml"
echo email = "">> "%USERPROFILE%\.streamlit\credentials.toml"
:cred_ok

rem --- Crea logo.ico (icono del acceso directo) desde Imgs\logo.png si aun no existe ---
if exist "Imgs\logo.png" if not exist "logo.ico" ".venv\Scripts\python.exe" -c "from PIL import Image; Image.open('Imgs/logo.png').convert('RGBA').save('logo.ico', sizes=[(16,16),(32,32),(48,48),(64,64),(128,128),(256,256)])"

rem --- Si la app ya esta corriendo, solo abre el navegador ---
netstat -ano | findstr /R /C:":8501 .*LISTENING" >nul
if not errorlevel 1 goto :ya_abierta

echo.
echo Abriendo la aplicacion en el navegador...
echo No cierres esta ventana mientras uses la app. Para apagarla, cierrala.
echo.
".venv\Scripts\python.exe" -m streamlit run app.py --server.port 8501 --browser.gatherUsageStats false
pause
exit /b

:ya_abierta
start "" http://localhost:8501
exit /b

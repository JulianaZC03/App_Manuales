@echo off
rem Crea en el Escritorio un acceso directo "Manuales OyM" que abre iniciar_app.bat.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$d=[Environment]::GetFolderPath('Desktop'); $s=(New-Object -ComObject WScript.Shell).CreateShortcut($d+'\Manuales OyM.lnk'); $s.TargetPath='%~dp0iniciar_app.bat'; $s.WorkingDirectory='%~dp0'; $s.IconLocation='%~dp0logo.ico'; $s.Description='Generador de manuales OyM'; $s.Save()"
echo.
echo Listo: busca "Manuales OyM" en el Escritorio.
pause

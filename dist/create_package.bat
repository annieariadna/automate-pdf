@echo off
echo 🎁 Creando paquete de distribución...

set PACKAGE_NAME=ExtractorPDF_v1.0
set PACKAGE_DIR=%PACKAGE_NAME%

if exist %PACKAGE_DIR% rmdir /s /q %PACKAGE_DIR%
mkdir %PACKAGE_DIR%

echo 📦 Copiando archivos...
copy /Y ExtractorPDF.exe %PACKAGE_DIR%\
copy /Y README.txt %PACKAGE_DIR%\
copy /Y test.bat %PACKAGE_DIR%\

mkdir %PACKAGE_DIR%\output

echo 🗜️ Creando archivo ZIP...
powershell Compress-Archive -Path %PACKAGE_DIR% -DestinationPath %PACKAGE_NAME%.zip -Force

echo ✅ Paquete creado: %PACKAGE_NAME%.zip
echo.
echo El paquete incluye:
echo   • ExtractorPDF.exe (ejecutable principal)
echo   • README.txt (instrucciones)
echo   • test.bat (script de prueba)
echo   • output\ (carpeta para archivos de salida)
echo.
pause

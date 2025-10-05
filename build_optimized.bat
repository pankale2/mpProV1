@echo off
echo Building optimized EXE...

REM Clean previous builds
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM Create virtual environment for clean build
python -m venv build_env
call build_env\Scripts\activate

REM Install minimal dependencies
pip install --no-cache-dir -r requirements.txt
pip install --no-cache-dir pyinstaller

REM Build with optimization
pyinstaller --clean --noconfirm RIDPIDProcessor.spec

REM Compress with UPX (if available)
if exist "C:\Program Files\upx\upx.exe" (
    echo Compressing with UPX...
    "C:\Program Files\upx\upx.exe" --best --lzma dist\RIDPIDProcessor.exe
)

REM Cleanup build environment
call deactivate
rmdir /s /q build_env

echo Build complete! EXE size:
for %%A in (dist\RIDPIDProcessor.exe) do echo %%~zA bytes
pause

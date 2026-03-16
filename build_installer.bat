@echo off
setlocal EnableDelayedExpansion

:: ============================================================
:: build_installer.bat
:: Builds the full OptimumK SolidWorks Plugin installer.
::
:: Prerequisites:
::   - .NET SDK / MSBuild (for the C# project)
::   - Python + PyInstaller  (pip install pyinstaller)
::   - Inno Setup 6          (https://jrsoftware.org/isdl.php)
:: ============================================================

set SCRIPT_DIR=%~dp0
set SRC_DIR=%SCRIPT_DIR%installer\source

:: ── Step 1: Build C# project ────────────────────────────────
echo.
echo [1/4] Building SuspensionTools (C#)...
dotnet build "%SCRIPT_DIR%sw_drawer\sw_drawer.csproj" -c Release --nologo
if errorlevel 1 (
    echo ERROR: C# build failed.
    exit /b 1
)

:: ── Step 2: Build Python exe via PyInstaller ────────────────
echo.
echo [2/4] Building Python exe (PyInstaller)...
pyinstaller "%SCRIPT_DIR%build.spec" --noconfirm
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    exit /b 1
)

:: ── Step 3: Collect files for the installer ──────────────────
echo.
echo [3/4] Staging files...
if not exist "%SRC_DIR%" mkdir "%SRC_DIR%"

:: Python exe
copy /Y "%SCRIPT_DIR%dist\OptimumK_SolidWorks_Plugin.exe" "%SRC_DIR%\" || goto :copy_err

:: C# exe
copy /Y "%SCRIPT_DIR%sw_drawer\bin\Release\net48\SuspensionTools.exe" "%SRC_DIR%\" || goto :copy_err

:: Part file (must exist in project root)
if not exist "%SCRIPT_DIR%Marker.SLDPRT" (
    echo ERROR: Marker.SLDPRT not found in project root.
    exit /b 1
)
copy /Y "%SCRIPT_DIR%Marker.SLDPRT" "%SRC_DIR%\" || goto :copy_err

:: Icon
copy /Y "%SCRIPT_DIR%icon.ico" "%SRC_DIR%\" || goto :copy_err

:: Help file and images
copy /Y "%SCRIPT_DIR%help.htm" "%SRC_DIR%\" || goto :copy_err
if not exist "%SRC_DIR%\help_files" mkdir "%SRC_DIR%\help_files"
xcopy /Y /Q "%SCRIPT_DIR%help_files\*.png" "%SRC_DIR%\help_files\" || goto :copy_err

:: ── Step 4: Compile Inno Setup installer ─────────────────────
echo.
echo [4/4] Compiling installer...

:: Try the default Inno Setup 6 install location
set ISCC="%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
if not exist %ISCC% set ISCC="%ProgramFiles%\Inno Setup 6\ISCC.exe"

if not exist %ISCC% (
    echo ERROR: ISCC.exe not found. Install Inno Setup 6 or add it to PATH.
    exit /b 1
)

%ISCC% "%SCRIPT_DIR%installer\setup.iss"
if errorlevel 1 (
    echo ERROR: Inno Setup compilation failed.
    exit /b 1
)

echo.
echo ============================================================
echo Done!  Installer: %SCRIPT_DIR%installer\OptimumK_Setup.exe
echo ============================================================
exit /b 0

:copy_err
echo ERROR: File copy failed.
exit /b 1

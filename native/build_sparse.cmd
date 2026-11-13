@echo off
setlocal
if "%~1"=="" exit /b 2
if not exist "%~1\Eigen\SparseCholesky" exit /b 2
for %%i in ("%~1") do set "EIGENROOT=%%~fi"
for /f "usebackq tokens=*" %%i in (`"%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do set "VSROOT=%%i"
if not defined VSROOT exit /b 1
call "%VSROOT%\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 exit /b 1
cd /d "%~dp0"
if not exist bin\sparse mkdir bin\sparse
cl /nologo /O2 /EHsc /std:c++17 /utf-8 /fp:precise /MT /LD /DNDEBUG /I "%EIGENROOT%" /Fo:bin\sparse\soilfem_sparse.obj soilfem_sparse.cpp /Fe:bin\sparse\soilfem_sparse_x64.dll /link /INCREMENTAL:NO
exit /b %errorlevel%

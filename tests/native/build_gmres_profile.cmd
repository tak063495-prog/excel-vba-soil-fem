@echo off
setlocal
if "%~2"=="" exit /b 1
for /f "usebackq tokens=*" %%i in (`"%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do set "VSROOT=%%i"
if not defined VSROOT exit /b 1
call "%VSROOT%\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 exit /b 1
if not exist "%~2" mkdir "%~2"
pushd "%~dp0..\..\native"
cl /nologo /O2 /fp:strict /MT /c /W4 /WX /std:c11 /I . /Fo"%~2\sfn_gmres.obj" "%~1"
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /LD /W4 /WX /std:c11 /Fo"%~2\soilfem_native.obj" soilfem_native.c bin\sfn_fp.obj bin\sfn_simd.obj bin\sfn_state.obj bin\sfn_assembly.obj "%~2\sfn_gmres.obj" /Fe"%~2\soilfem_native_x64.dll" /link /INCREMENTAL:NO
exit /b %errorlevel%

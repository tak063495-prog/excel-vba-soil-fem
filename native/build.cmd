@echo off
setlocal
for /f "usebackq tokens=*" %%i in (`"%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do set "VSROOT=%%i"
if not defined VSROOT exit /b 1
call "%VSROOT%\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 exit /b 1
cd /d "%~dp0"
if not exist bin mkdir bin
ml64 /nologo /c /Fo bin\sfn_fp.obj sfn_fp.asm
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:precise /arch:AVX2 /MT /c /W4 /WX /std:c11 /Fo:bin\sfn_simd.obj sfn_simd.c
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /c /W4 /WX /std:c11 /Fo:bin\sfn_state.obj sfn_state.c
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /c /W4 /WX /std:c11 /Fo:bin\sfn_gmres.obj sfn_gmres.c
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /LD /W4 /WX /std:c11 /Fo:bin\ soilfem_native.c bin\sfn_fp.obj bin\sfn_simd.obj bin\sfn_state.obj bin\sfn_gmres.obj /Fe:bin\soilfem_native_x64.dll /link /INCREMENTAL:NO
exit /b %errorlevel%

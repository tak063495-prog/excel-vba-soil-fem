@echo off
setlocal
call "C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 exit /b 1
cd /d "%~dp0"
if not exist bin mkdir bin
ml64 /nologo /c /Fo bin\sfn_fp.obj sfn_fp.asm
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /LD /W4 /WX /std:c11 /Fo:bin\ soilfem_native.c bin\sfn_fp.obj /Fe:bin\soilfem_native_x64.dll /link /INCREMENTAL:NO
exit /b %errorlevel%

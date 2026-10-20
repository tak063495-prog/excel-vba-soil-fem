@echo off
setlocal
rem Build one DLL containing both the MSVC native solver exports (sfn_*) and
rem the GCC/x87 material exports (sfm_*).  The material object is compiled
rem with the exact flags used by build_mat.sh; it is never recompiled with
rem MSVC, because its x87 inline assembly and MSVCR100 binding are part of
rem the numerical contract.
for /f "usebackq tokens=*" %%i in (`"%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe" -latest -products * -requires Microsoft.VisualStudio.Component.VC.Tools.x86.x64 -property installationPath`) do set "VSROOT=%%i"
if not defined VSROOT exit /b 1
call "%VSROOT%\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 exit /b 1
where x86_64-w64-mingw32-gcc >nul 2>nul
if errorlevel 1 (
  if exist "%~dp0..\tests\tmp\toolchain14\extracted\w64devkit\bin\x86_64-w64-mingw32-gcc.exe" set "PATH=%~dp0..\tests\tmp\toolchain14\extracted\w64devkit\bin;%PATH%"
)
where x86_64-w64-mingw32-gcc >nul 2>nul
if errorlevel 1 (
  echo x86_64-w64-mingw32-gcc is required to preserve soilfem_mat.c x87 semantics.
  exit /b 2
)
for /f "usebackq tokens=*" %%i in (`x86_64-w64-mingw32-gcc -print-libgcc-file-name`) do set "LIBGCC=%%i"
if not defined LIBGCC (
  echo Could not locate libgcc.a.
  exit /b 2
)
cd /d "%~dp0"
if not exist bin\combined mkdir bin\combined
x86_64-w64-mingw32-gcc -O3 -std=c11 -Wall -Wextra -Werror -Wno-unused-parameter -Wno-misleading-indentation -mfpmath=387 -fexcess-precision=standard -ffp-contract=off -fno-fast-math -fno-builtin -fno-stack-protector -mno-stack-arg-probe -Dmemcpy=sfm_memcpy -Dmemset=sfm_memset -c soilfem_mat.c -o bin\combined\soilfem_mat.o
if errorlevel 1 exit /b 1
ml64 /nologo /c /Fo bin\combined\sfn_fp.obj sfn_fp.asm
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:precise /arch:AVX2 /MT /c /W4 /WX /std:c11 /Fo:bin\combined\sfn_simd.obj sfn_simd.c
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /c /W4 /WX /std:c11 /Fo:bin\combined\sfn_state.obj sfn_state.c
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /c /W4 /WX /std:c11 /Fo:bin\combined\sfn_gmres.obj sfn_gmres.c
if errorlevel 1 exit /b 1
cl /nologo /O2 /fp:strict /MT /c /W4 /WX /std:c11 /Fo:bin\combined\sfn_assembly.obj sfn_assembly.c
if errorlevel 1 exit /b 1
rem link.exe accepts the MinGW COFF object. libgcc is supplied explicitly;
rem omitting it can silently change the generated ABI or leave unresolved
rem helper calls on a different compiler version.
cl /nologo /O2 /fp:strict /MT /LD /W4 /WX /std:c11 /Fo:bin\combined\soilfem_native.obj soilfem_native.c bin\combined\soilfem_mat.o bin\combined\sfn_fp.obj bin\combined\sfn_simd.obj bin\combined\sfn_state.obj bin\combined\sfn_gmres.obj bin\combined\sfn_assembly.obj "%LIBGCC%" /Fe:bin\combined\soilfem_native_x64.dll /link /DEF:soilfem_combined.def /INCREMENTAL:NO
exit /b %errorlevel%

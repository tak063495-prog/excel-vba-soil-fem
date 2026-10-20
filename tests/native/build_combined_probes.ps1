param([string]$OutputDirectory='tests/tmp/combined_probe_repro',[int]$Count=100000)
$ErrorActionPreference='Stop';$repo=Split-Path -Parent (Split-Path -Parent $PSScriptRoot);Push-Location $repo
try {
  if(Test-Path -LiteralPath $OutputDirectory){throw 'Choose a new output directory; prior probe evidence is preserved.'}
  $out=(New-Item -ItemType Directory -Path $OutputDirectory).FullName
  $gcc=Get-Command x86_64-w64-mingw32-gcc -ErrorAction SilentlyContinue
  if(-not $gcc){$fallback=Join-Path $repo 'tests/tmp/toolchain14/extracted/w64devkit/bin';$candidate=Join-Path $fallback 'x86_64-w64-mingw32-gcc.exe';if(Test-Path $candidate){$env:Path="$fallback;$env:Path";$gcc=Get-Command x86_64-w64-mingw32-gcc}}
  if(-not $gcc){throw 'x86_64-w64-mingw32-gcc not found'}
  python tests/native/test_material_points.py --prepare (Join-Path $out 'material_probe.c')
  if($LASTEXITCODE -ne 0){throw 'Probe source preparation failed'}
  $flags=@('-O3','-std=c11','-Wall','-Wextra','-Werror','-Wno-unused-parameter','-Wno-misleading-indentation','-mfpmath=387','-fexcess-precision=standard','-ffp-contract=off','-fno-fast-math','-fno-builtin','-fno-stack-protector','-mno-stack-arg-probe')
  & $gcc.Source @flags -shared '-Wl,--entry=DllMain,--kill-at' -s -o "$out/soilfem_mat_probe.dll" "$out/material_probe.c" -lkernel32 -lgcc
  if($LASTEXITCODE -ne 0){throw 'Reference probe build failed'}
  & $gcc.Source @flags -Dmemcpy=sfm_memcpy -Dmemset=sfm_memset -c "$out/material_probe.c" -o "$out/material_probe.o"
  if($LASTEXITCODE -ne 0){throw 'Combined material object build failed'}
  Copy-Item native/soilfem_combined.def (Join-Path $out 'combined_probe.def');Add-Content (Join-Path $out 'combined_probe.def') "    sfm_probe_size`n    sfm_probe_point"
  $vs=(Get-ChildItem 'C:\Program Files\Microsoft Visual Studio' -Recurse -Filter vcvars64.bat -ErrorAction SilentlyContinue|Select-Object -First 1).FullName;if(-not $vs){throw 'vcvars64.bat not found'}
  foreach($name in 'sfn_fp','sfn_simd','sfn_state','sfn_gmres','sfn_assembly'){if(-not(Test-Path "native/bin/combined/$name.obj")){throw 'Build production objects with native/build_combined.cmd first'}}
  $libgcc=(& $gcc.Source -print-libgcc-file-name).Trim()
  $cmd='call "'+$vs+'" >nul && cl /nologo /O2 /fp:strict /MT /LD /W4 /WX /std:c11 /Fo:"'+$out+'\soilfem_native_probe.obj" native\soilfem_native.c "'+$out+'\material_probe.o" native\bin\combined\sfn_fp.obj native\bin\combined\sfn_simd.obj native\bin\combined\sfn_state.obj native\bin\combined\sfn_gmres.obj native\bin\combined\sfn_assembly.obj "'+$libgcc+'" /Fe:"'+$out+'\soilfem_native_probe.dll" /link /DEF:"'+$out+'\combined_probe.def" /INCREMENTAL:NO'
  cmd /c $cmd
  if($LASTEXITCODE -ne 0){throw 'Combined probe link failed'}
  python tests/native/test_material_points.py --reference (Join-Path $out 'soilfem_mat_probe.dll') --candidate (Join-Path $out 'soilfem_native_probe.dll') --count $Count --out (Join-Path $out 'material_compare.json')
  if($LASTEXITCODE -ne 0){throw 'Material point regression failed'}
  Write-Output "PASS probes: $out"
} finally {Pop-Location}

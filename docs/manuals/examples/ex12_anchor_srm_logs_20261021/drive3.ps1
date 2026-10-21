$d='C:\Users\link_\manual_shots_20261021\run9'
function L($m){ "$(Get-Date -f HH:mm:ss) $m"|Out-File -Append "$d\main3.log" -Encoding utf8 }
cd "$env:USERPROFILE\manual_rewrite_20261021"
L 'start anchor_inco_srm'
powershell -NoProfile -ExecutionPolicy Bypass -File "$d\run_manual_case_bom.ps1" -Fixture tests/manual_cases/slope_wall_mesh.json -OutDirectory "$d\anchor_inco_srm_2204" -Anchor -Srm -Flow INCONSISTENT *>> "$d\anchor_inco_srm.log"
L "done $LASTEXITCODE"

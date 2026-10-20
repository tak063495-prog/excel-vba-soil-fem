[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)][string]$Book,
  [string]$Out = "tests/results/workbook_inputs.json"
)
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
Push-Location $root
try {
  $bookPath = (Resolve-Path -LiteralPath $Book).Path
  if (-not (Test-Path -LiteralPath $bookPath -PathType Leaf)) { throw "Book not found: $Book" }
  $running = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue)
  if ($running.Count -gt 0) { throw "Excel is already running; close it before this owned-workbook test." }
  $tmpDir = Join-Path $root 'tests/tmp/workbook_inputs'
  New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null
  $copy = Join-Path $tmpDir ([IO.Path]::GetFileNameWithoutExtension($bookPath) + '_test.xlsm')
  Copy-Item -LiteralPath $bookPath -Destination $copy -Force
  $checks = [Collections.Generic.List[object]]::new()
  function Check([bool]$ok,[string]$name,[string]$detail='') {
    if (-not $ok) { throw "FAIL $name $detail" }
    $checks.Add([ordered]@{name=$name;status='PASS';detail=$detail})
  }
  function ReadCell($ws,$addr) { return $ws.Range($addr).Value2 }
  $xl = $null; $wb = $null; $h = $null
  try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false; $xl.EnableEvents = $false; $xl.AutomationSecurity = 1
    $wb = $xl.Workbooks.Open($copy,0,$false)
    $h = $wb.VBProject.VBComponents.Add(1); $h.Name = 'WorkbookInputHarness'
    $h.CodeModule.AddFromString(@'
Option Explicit
Public Function TValidateSettings() As String
  On Error GoTo Bad
  Call FEMUiValidateSettings: TValidateSettings = "PASS": Exit Function
Bad: TValidateSettings = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Function TValidateModel() As String
  On Error GoTo Bad
  Call FEMInvalidateInputCache: NumberOfNode = SetNumberOfNode(): Call ValidateInputData: TValidateModel = "PASS": Exit Function
Bad: TValidateModel = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Function TGeometry() As String
  On Error GoTo Bad
  Call FEMUiBuildGeometry: TGeometry = "PASS": Exit Function
Bad: TGeometry = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Function TRead(ByVal keyName As String, ByVal isText As Boolean) As Variant
  On Error GoTo Bad
  If isText Then TRead = FEMReadTextSetting(keyName, "") Else TRead = FEMReadSetting(keyName, -999999#)
  Exit Function
Bad: TRead = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Sub TSetCellError(ByVal pageName As String, ByVal cellAddress As String)
  ThisWorkbook.Worksheets(pageName).Range(cellAddress).Value2 = CVErr(xlErrNA)
End Sub
'@)
    $prefix = "'" + $wb.Name + "'!"
    $xl.Run($prefix + 'FEMWriteBuildStamp', $false)
    $xl.CalculateFull()
    Check (($xl.Run($prefix+'TValidateSettings') -eq 'PASS')) 'recommended settings validate'

    $pages = @('メッシュ設定','解析設定','高速化設定','出力・表示','詳細設定')
    $keys = @{}
    $legacy = $wb.Worksheets.Item('設定'); $legacyRows=@{}
    $legacyLast=$legacy.Cells($legacy.Rows.Count,5).End(-4162).Row
    $registryData=$legacy.Range("C1:E$legacyLast").Value2
    for($lr=1;$lr -le $legacyLast;$lr++){$rk=[string]$registryData.GetValue($lr,3);if($rk){$legacyRows[$rk]=$lr}}
    foreach ($page in $pages) {
      $ws = $wb.Worksheets.Item($page); $last = $ws.Cells($ws.Rows.Count,6).End(-4162).Row
      for ($r=1; $r -le $last; $r++) {
        $key = [string]$ws.Cells($r,6).Value2
        if ($key.Trim().Length -eq 0) { continue }
        Check (-not $keys.ContainsKey($key.ToUpperInvariant())) 'UI keys unique' "$page!F$r=$key"
        $keys[$key.ToUpperInvariant()] = "$page!C$r"
        Check ([string]$ws.Cells($r,7).Value2 -ne '') 'recommended help present' "$page!$key"
        Check ($legacyRows.ContainsKey($key)) 'legacy key exists' $key
        $lr=$legacyRows[$key]
        Check ($legacy.Cells($lr,3).HasFormula) 'legacy binding formula' $key
        $expected=$page+"!C"+$r
        Check (([string]$legacy.Cells($lr,3).Formula).Replace("'",'').Contains($expected)) 'legacy formula targets exact input' $key
      }
    }
    foreach ($k in @('NATIVE_DLL','NATIVE_PRECISION','NATIVE_STATE_COPY','NATIVE_ASSEMBLY','NATIVE_GMRES_SIMD','NATIVE_GMRES_TUNED','NATIVE_GMRES_PREPARED','IO_FAST_PATH','EVENT_LOG_FAST')) { Check $keys.ContainsKey($k) 'native key exposed' $k }
    Check (($xl.Run($prefix+'TRead','NATIVE_DLL',$true) -eq 'AUTO')) 'native DLL recommended AUTO'
    Check (($xl.Run($prefix+'TRead','NATIVE_PRECISION',$true) -eq 'AUTO')) 'native precision recommended AUTO'
    foreach ($flag in @('NATIVE_STATE_COPY','NATIVE_ASSEMBLY','NATIVE_GMRES_SIMD','NATIVE_GMRES_TUNED','NATIVE_GMRES_PREPARED','IO_FAST_PATH','EVENT_LOG_FAST')) { Check ([double]($xl.Run($prefix+'TRead',$flag,$false)) -eq 1) 'native flag recommended 1' $flag }

    foreach ($pair in @(@('FLOW_POLICY','INCONSISTENT'),@('FLOW_POLICY','DAVIS'),@('RCM_POLICY','AUTO'),@('RCM_POLICY','ON'),@('RCM_POLICY','OFF'),@('NATIVE_DLL','AUTO'),@('NATIVE_DLL','BAND'),@('NATIVE_DLL','OFF'),@('NATIVE_PRECISION','STRICT'),@('NATIVE_PRECISION','SSE2'),@('NATIVE_PRECISION','AUTO'),@('NATIVE_PRECISION','AVX2'))) {
      $key=$pair[0]; $cell=$keys[$key]; $ws=$wb.Worksheets.Item(($cell -split '!')[0]); $addr=($cell -split '!')[1]; $old=$ws.Range($addr).Value2; $ws.Range($addr).Value2=$pair[1]; Check (($xl.Run($prefix+'TValidateSettings') -eq 'PASS')) 'enum accepted' "$key=$($pair[1])"; $ws.Range($addr).Value2=$old
    }
    $nativeCell = $keys['NATIVE_DLL']; $nws=$wb.Worksheets.Item(($nativeCell -split '!')[0]); $na=($nativeCell -split '!')[1]; $old=$nws.Range($na).Value2; $nws.Range($na).Value2='ON'; $nativeRejectResult=[string]$xl.Run($prefix+'TValidateSettings'); Check ($nativeRejectResult.StartsWith('ERROR|')) 'invalid native enum rejected' ($nativeRejectResult+' value='+[string]$nws.Range($na).Value2+' key='+$nativeCell); $nws.Range($na).Value2=$old

    $numericCases = @(
      @('MESH_MATERIAL',1,$true),@('MESH_MATERIAL',0,$false),@('MESH_MATERIAL',1.5,$false),@('MESH_MATERIAL','',$false),@('MESH_MATERIAL',2147483648,$false),
      @('Q8_HOURGLASS_FACTOR',0,$true),@('Q8_HOURGLASS_FACTOR',1,$true),@('Q8_HOURGLASS_FACTOR',-1,$false),@('Q8_HOURGLASS_FACTOR',1.1,$false),
      @('SRM_TOL',0,$false),@('SRM_TOL',0.0125,$true),@('VIEW_DISP_SCALE',0,$true),@('DEBUG_FLUSH',0,$true),
      @('VIEW_MIN_X',-2.5,$true),@('VIEW_MAX_Y',-3,$true),@('NATIVE_STATE_COPY',0,$true),@('NATIVE_STATE_COPY',1,$true),@('NATIVE_STATE_COPY',0.5,$false),@('MESH_MATERIAL','1',$false),@('VIEW_COLOR_RESULT',99999,$false),@('MESH_REPAIR_TERRAIN_COMPARE',0.5,$false),@('MESH_REPAIR_TERRAIN_COMPARE',0,$true),@('MESH_REPAIR_TERRAIN_COMPARE',1,$true)
    )
    foreach ($case in $numericCases) {
      $key=$case[0]; $cell=$keys[$key]; $ws=$wb.Worksheets.Item(($cell -split '!')[0]); $addr=($cell -split '!')[1]; $oldFormula=$ws.Range($addr).Formula; $oldFormat=$ws.Range($addr).NumberFormat; if($case[1] -is [string]){$ws.Range($addr).NumberFormat='@'}; if($case[1] -is [ValueType]){$ws.Range($addr).Value2=[double]$case[1]}else{$ws.Range($addr).Value2=[string]$case[1]}
      $validation=$false; try { $validation=[bool]$ws.Range($addr).Validation.Value } catch { $validation=$false }
      $result=[string]$xl.Run($prefix+'TValidateSettings'); $accepted=($result -eq 'PASS')
      Check (($validation -eq [bool]$case[2]) -and ($accepted -eq [bool]$case[2])) 'numeric validation and runtime agree' "$key=$($case[1]) result=$result validation=$validation"
      $ws.Range($addr).NumberFormat=$oldFormat; if($oldFormula -is [ValueType]){$ws.Range($addr).Value2=[double]$oldFormula}else{$ws.Range($addr).Formula=[string]$oldFormula}
    }
    $q8=$keys['Q8_HOURGLASS_FACTOR']; $qws=$wb.Worksheets.Item(($q8 -split '!')[0]); $qa=($q8 -split '!')[1]; $qold=$qws.Range($qa).Formula; $qws.Range($qa).Formula='=NA()'; Check (($xl.Run($prefix+'TValidateSettings') -like 'ERROR|*')) 'formula error rejected'; $qws.Range($qa).Formula=$qold

    $resumeInfo=$keys['EXPORT_LOAD']; $rw=$wb.Worksheets.Item(($resumeInfo -split '!')[0]); $ra=($resumeInfo -split '!')[1]
    $rv=[string]$rw.Range($ra).Value2
    $xl.Run($prefix+'TSetCellError',($resumeInfo -split '!')[0],$ra)
    Check (($xl.Run($prefix+'TValidateSettings') -like '*EXPORT_LOAD*')) 'resume error value rejected with key'
    $xl.Run($prefix+'FEMWriteTextSetting','EXPORT_LOAD',$rv)
    $xl.EnableEvents=$false; $xl.Calculation=-4135
    $before = @{}; foreach ($k in $keys.Keys) { $c=$keys[$k]; $w=$wb.Worksheets.Item(($c -split '!')[0]); $before[$k]=$w.Range(($c -split '!')[1]).Value2 }
    $xl.Run($prefix+'FEMWriteBuildStamp',$false); $xl.Run($prefix+'TValidateSettings') | Out-Null
    foreach ($k in $before.Keys) { $c=$keys[$k]; $w=$wb.Worksheets.Item(($c -split '!')[0]); Check ([string]$w.Range(($c -split '!')[1]).Value2 -eq [string]$before[$k]) 'layout preserves setting' $k }
    # A dependent setting formula can have a valid stale value in manual mode.
    $helper=$wb.Worksheets.Item('操作パネル').Range('Z1');$helper.Value2=0.5
    $q8=$keys['Q8_HOURGLASS_FACTOR'];$qws=$wb.Worksheets.Item(($q8 -split '!')[0]);$qa=($q8 -split '!')[1];$qold=$qws.Range($qa).Value2
    $qws.Range($qa).Formula="='操作パネル'!Z1";$qws.Calculate()
    Check ($qws.Range($qa).Value2 -eq 0.5) 'manual formula begins valid'
    $helper.Value2=2.0
    Check ($qws.Range($qa).Value2 -eq 0.5) 'manual setting formula has stale cached value before validation'
    Check (($xl.Run($prefix+'TValidateSettings') -like '*Q8_HOURGLASS_FACTOR*')) 'validation recalculates manual formula and rejects new invalid value'
    $qws.Range($qa).Value2=[double]$qold;$helper.ClearContents()
    Check (($xl.Run($prefix+'TValidateModel') -eq 'PASS')) 'recommended model validates'
    $modelCases = @(@('材料データ','B2',0),@('材料データ','C2',0.5),@('要素データ','B2',999999),@('節点データ','A3',$null),@('節点データ','D2',0.5))
    foreach ($mc in $modelCases) {
      $mws=$wb.Worksheets.Item($mc[0]); $mcell=$mws.Range($mc[1]); $mold=$mcell.Value2
      if ($mc[0] -eq '節点データ' -and $mc[1] -eq 'A3') { $mcell.Value2=[double]$mws.Range('A2').Value2 } else { $mcell.Value2=[double]$mc[2] }
      Check (($xl.Run($prefix+'TValidateModel') -like 'ERROR|*')) 'invalid model input rejected' "$($mc[0])!$($mc[1])"
      $mcell.Value2=[double]$mold
    }
    $shape=$wb.Worksheets.Item('形状入力'); $legacyShape=$wb.Worksheets.Item('図形定義'); $shapeOld=$shape.Range('B9').Value2; $shapeBefore=($legacyShape.Range('A1:N200').Value2 | ConvertTo-Json -Compress -Depth 4); $shape.Range('B9').Value2=[double]$shape.Range('B8').Value2; Check (($xl.Run($prefix+'TGeometry') -like 'ERROR|*')) 'duplicate geometry rejected'; $shape.Range('B9').Value2=[double]$shapeOld; $shapeAfter=($legacyShape.Range('A1:N200').Value2 | ConvertTo-Json -Compress -Depth 4); Check ($shapeBefore -eq $shapeAfter) 'geometry legacy data unchanged after rejected input'
    foreach ($page in $pages) {
      $s=$wb.Worksheets.Item($page); $last=$s.Cells($s.Rows.Count,6).End(-4162).Row
      for ($r=1; $r -le $last; $r++) { $key=[string]$s.Cells($r,6).Value2; if ($key.Trim().Length -eq 0) { continue }; Check ($s.Cells($r,3).Locked -eq $key.StartsWith('MESH_GEN_LAST_')) 'input or automatic output lock' "$page!C$r"; Check ($s.Cells($r,6).Locked -and $s.Cells($r,7).Locked) 'setting metadata locked' "$page!F$r/G$r"; if($key -notmatch '^(MESH_GEN_LAST_|BAND_LU_BENCH|EXPORT_LOAD$)'){Check ($s.Cells($r,3).Validation.AlertStyle -eq 1) 'validation stop style' $key} }
    }
    foreach ($addr in @('B101')) { Check ($wb.Worksheets.Item('ステージ').Range($addr).Validation.AlertStyle -eq 1) 'stage validation stop' $addr }
    Check ($wb.Worksheets.Item('材料データ').Range('K101').Validation.AlertStyle -eq 1) 'material validation stop K101' 'K101'
    Check ($wb.Worksheets.Item('接合').Range('D101').Validation.AlertStyle -eq 1) 'joint validation stop D101' 'D101'

    $outPath = Join-Path $root $Out; New-Item -ItemType Directory -Force -Path (Split-Path $outPath) | Out-Null
    [ordered]@{book=$bookPath;copy=$copy;checks=$checks;count=$checks.Count;status='PASS'} | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $outPath -Encoding UTF8
    Write-Output ("PASS workbook input checks=" + $checks.Count + " output=" + $outPath)
  } finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
    if ($null -ne $xl) { try { $xl.EnableEvents=$false; $xl.Quit() } catch {} }
    if ($null -ne $wb) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($wb) | Out-Null }
    if ($null -ne $xl) { [Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl) | Out-Null }
  }
} finally { Pop-Location }







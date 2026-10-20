param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$OutDirectory,[switch]$UseDisplaySource)
$ErrorActionPreference='Stop';$repo=Split-Path -Parent $PSScriptRoot;Push-Location $repo
$app=$null;$wb=$null;$checks=[Collections.Generic.List[string]]::new()
function Check($ok,[string]$label){if(-not $ok){throw "FAIL $label"};$checks.Add($label)}
try{
 if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
 $Book=(Resolve-Path -LiteralPath $Book).Path
 $OutDirectory=[IO.Path]::GetFullPath($OutDirectory);New-Item -ItemType Directory -Force $OutDirectory|Out-Null
 $tmp=Join-Path $repo ('tests/tmp/increment_results_'+[Guid]::NewGuid().ToString('N'));New-Item -ItemType Directory $tmp|Out-Null
 $copy=Join-Path $tmp 'results.xlsm';Copy-Item -LiteralPath $Book -Destination $copy
 $app=New-Object -ComObject Excel.Application;$app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1
 $wb=$app.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!"
 $displaySources=@{}
 if($UseDisplaySource){
  foreach($name in 'FEMIncrementSvg','FEMResultOptions'){
   $path=Join-Path $repo ('src/vba/'+$name+'.bas');$module=$wb.VBProject.VBComponents.Item($name).CodeModule
   $module.DeleteLines(1,$module.CountOfLines);$module.AddFromString((Get-Content -LiteralPath $path -Raw -Encoding UTF8))
   $displaySources[$name]=(Get-FileHash -LiteralPath $path).Hash
  }
 }
 $h=$wb.VBProject.VBComponents.Add(1);$h.Name='IncrementResultTests'
 $h.CodeModule.AddFromString(@'
Public Function TDrawIncrement() As String
 SuppressUserMessages = True
 FEMViewResult
 TDrawIncrement = FEMViewerLastError
End Function
Public Function TReadIncrement() As String
 On Error GoTo Bad
 Dim d As Object
 Set d = FEMIncrementRead(2)
 TReadIncrement = "PASS": Exit Function
Bad:
 TReadIncrement = "ERROR|" & Err.Description
End Function
'@)
 $compile=$app.VBE.CommandBars.FindControl(1,578);if($compile.Enabled){$compile.Execute()}
 Check (-not $app.VBE.CommandBars.FindControl(1,578).Enabled) 'Excel full compile with result harness'
 if($UseDisplaySource){$app.Run($prefix+'FEMResultOptionsInstall')}
 $inc=$wb.Worksheets.Item('結果増分');$last=$inc.Cells($inc.Rows.Count,1).End(-4162).Row
 Check ($last -eq 81) '20 continuum elements x four GP saved'
 $data=$inc.Range("A2:T$last").Value2
 $rows=@();for($r=1;$r -le $data.GetLength(0);$r++){$a=@();for($c=1;$c -le 20;$c++){$a+=,$data.GetValue($r,$c)};$rows+=,@($a)}
 ConvertTo-Json -InputObject $rows -Depth 4|Set-Content (Join-Path $OutDirectory 'increment_rows.json') -Encoding UTF8
 $ws=$wb.Worksheets.Item('出力・表示');$ws.Range('C18').Value2=[string]$app.Run($prefix+'FEMResultChoiceLabel',2)
 $app.EnableEvents=$true
 foreach($mode in 'A','B'){
  $ws.Range('C27').Value2=if($mode -eq 'A'){'A：選択結果へ等値線を追加'}else{'B：増分専用カラー'}
  foreach($s in $false,$true){foreach($p in $false,$true){
   $ws.CheckBoxes('UI_ShearIncrement').Value=if($s){1}else{-4146}
   $ws.CheckBoxes('UI_PlasticIncrement').Value=if($p){1}else{-4146}
   Check ([string]$app.Run($prefix+'TDrawIncrement') -ceq '') "actual SRM result $mode shear=$s cross=$p draws"
   Check ($wb.Worksheets.Item('図').Shapes.Item('P1Agg_ViewSVG').AlternativeText -like '*20要素*') 'embedded SVG picture exists'
   Check ($wb.Names('_FEMViewerResultsDirty').RefersTo -ceq '=FALSE') 'display controls preserve clean calculation'
   if($s -and $p){$wb.SaveCopyAs((Join-Path $OutDirectory ($mode+'.xlsm')))}
  }}
 }
 foreach($id in 69,70){$ws.Range('C18').Value2=[string]$app.Run($prefix+'FEMResultChoiceLabel',$id);Check ([string]$app.Run($prefix+'TDrawIncrement') -ceq '') "increment content $id draws"}
 $wb.Save();$wb.Close($false);$app.EnableEvents=$true;$wb=$app.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!";$ws=$wb.Worksheets.Item('出力・表示');$inc=$wb.Worksheets.Item('結果増分')
 Check ([string]$app.Run($prefix+'TDrawIncrement') -ceq '') 'saved/reopened increment result draws without reanalysis'
 $app.EnableEvents=$false
 $old=$inc.Cells(2,3).Value2;$inc.Cells(2,3).Value2=1.5
 Check ([string]$app.Run($prefix+'TDrawIncrement') -like '*再解析*') 'fractional Gauss point rejected'
 $inc.Cells(2,3).Value2=[double]$old
 $old=$inc.Cells(6,18).Value2;$inc.Cells(6,18).Value2=[double]$old+0.1
 Check ([string]$app.Run($prefix+'TDrawIncrement') -like '*再解析*') 'mixed Fs across elements rejected'
 $inc.Cells(6,18).Value2=[double]$old
 $old=$inc.Cells(2,13).Value2;$inc.Cells(2,13).Value2=-0.001
 Check ([string]$app.Run($prefix+'TDrawIncrement') -like '*再解析*') 'negative maximum shear rejected'
 $inc.Cells(2,13).Value2=[double]$old
 $inc.Range('V1').Value2='OLD'
 Check ([string]$app.Run($prefix+'TDrawIncrement') -like '*最初から再解析*') 'missing old increment refuses fabricated display'
 $ws.Range('C18').Value2=[string]$app.Run($prefix+'FEMResultChoiceLabel',2);$ws.Range('J28').Value2=$false;$ws.Range('J29').Value2=$false
 Check ([string]$app.Run($prefix+'TDrawIncrement') -ceq '') 'legacy saved result displays with increments off'
 @{status='PASS';checks=$checks.ToArray();count=$checks.Count;source=$Book;display_sources=$displaySources}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $OutDirectory 'display_results.json') -Encoding UTF8
 Write-Output "PASS $($checks.Count) actual result display checks"
}finally{
 if($wb){try{$wb.Close($false)}catch{}}
 if($app){$app.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers();Pop-Location
}

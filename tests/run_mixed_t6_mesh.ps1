param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out,[switch]$Visible)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
$Book=(Resolve-Path -LiteralPath $Book).Path;$before=(Get-FileHash $Book).Hash;$Out=[IO.Path]::GetFullPath($Out)
if(Test-Path -LiteralPath $Out){throw 'Preserve previous evidence.'}
New-Item -ItemType Directory -Path $Out | Out-Null
$copy=Join-Path $Out 'mixed_t6.xlsm';Copy-Item $Book $copy
Copy-Item (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') $Out
$checks=[Collections.Generic.List[object]]::new();$xl=$null;$wb=$null
function Assert($ok,[string]$label){if(-not $ok){throw "FAIL: $label"};$checks.Add(@{label=$label;status='PASS'});"PASS $label"}
try{
 $xl=New-Object -ComObject Excel.Application;$xl.Visible=[bool]$Visible;$xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false)
 $mod=$wb.VBProject.VBComponents.Add(1);$mod.Name='MixedT6Driver';$mod.CodeModule.AddFromString([IO.File]::ReadAllText((Join-Path $PSScriptRoot 'mixed_t6_mesh_driver.bas')))
 function Raw([string]$name,[object[]]$parameters=@()){
  $macro="'"+$wb.Name+"'!"+$name
  switch($parameters.Count){0{$xl.Run($macro)} 1{$xl.Run($macro,$parameters[0])} 2{$xl.Run($macro,$parameters[0],$parameters[1])} 3{$xl.Run($macro,$parameters[0],$parameters[1],$parameters[2])} default{throw 'Too many parameters'}}
 }
 function Run([string]$name,[object[]]$parameters=@()){$v=[string](Raw $name $parameters);if($v.StartsWith('ERROR|')){throw $v};$v}
 function Snapshot([string]$sheet,[int]$columns=11){$w=$wb.Worksheets.Item($sheet);$last=$w.Cells($w.Rows.Count,1).End(-4162).Row;ConvertTo-Json -InputObject $w.Range($w.Cells(2,1),$w.Cells($last,$columns)).Value2 -Compress -Depth 6}
 function LoadFixture{
  $xl.EnableEvents=$false
  $j=Get-Content (Join-Path $PSScriptRoot 'fixtures/mixed_t6/q8_gravity.json') -Raw -Encoding UTF8|ConvertFrom-Json
  foreach($name in '節点データ','要素データ','材料データ','ステージ','載荷','接合'){$wb.Worksheets.Item($name).Range('A2:N20000').ClearContents()}
  foreach($pair in @(@('節点データ',$j.nodes),@('要素データ',$j.elements),@('材料データ',$j.materials),@('ステージ',$j.stages))){
   $w=$wb.Worksheets.Item($pair[0]);$rows=@($pair[1]);$cols=($rows|ForEach-Object {$_.Count}|Measure-Object -Maximum).Maximum
   $matrix=New-Object 'object[,]' $rows.Count,$cols
   for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt $rows[$r].Count;$c++){
    $v=$rows[$r][$c];if($null -ne $v){$matrix[$r,$c]=$(if($v -is [string]){[string]$v}else{[double]$v})}
   }}
   $w.Range($w.Cells(2,1),$w.Cells($rows.Count+1,$cols)).Value2=$matrix
  }
  $w=$wb.Worksheets.Item('節点データ');$e=$wb.Worksheets.Item('要素データ');$right=$j.elements[1]
  $w.Cells(([int]$right[3]+1),2).Value2=3.0
  $w.Cells(([int]$right[3]+1),3).Value2=0.05
  $w.Cells(([int]$right[4]+1),2).Value2=2.0
  foreach($row in $j.elements){for($k=0;$k -lt 4;$k++){
   $a=[int]$row[$k+1]+1;$b=[int]$row[(($k+1)%4)+1]+1;$m=[int]$row[$k+5]+1
   $w.Cells($m,2).Value2=([double]$w.Cells($a,2).Value2+[double]$w.Cells($b,2).Value2)/2
   $w.Cells($m,3).Value2=([double]$w.Cells($a,3).Value2+[double]$w.Cells($b,3).Value2)/2
  }}
  $wb.Worksheets.Item('水条件').Range('B3').Value2='NONE';$wb.Worksheets.Item('水条件').Range('A216:F415').ClearContents()
  $null=$xl.Run("'"+$wb.Name+"'!FEMWriteNumericSetting",'MESH_QUALITY_MIN_JACOBIAN_RATIO',0.000001)
  $null=$xl.Run("'"+$wb.Name+"'!FEMWriteNumericSetting",'MESH_QUALITY_MIN_ANGLE_DEG',10.0)
  $null=$xl.Run("'"+$wb.Name+"'!FEMWriteNumericSetting",'MESH_QUALITY_MAX_ANGLE_DEG',170.0)
  $null=$xl.Run("'"+$wb.Name+"'!FEMWriteNumericSetting",'MESH_QUALITY_MAX_ASPECT_RATIO',5.0)
  $null=$xl.Run("'"+$wb.Name+"'!FEMWriteTextSetting",'EXPORT_MODE','OFF')
  $null=$xl.Run("'"+$wb.Name+"'!FEMInvalidateRemeshedModel")
 }
 Assert ($wb.Worksheets.Item("メッシュ設定").Range("F30").Value2 -eq "MESH_T6_FALLBACK" -and $wb.Worksheets.Item("メッシュ設定").Range("C30").Value2 -eq 0) "release contains fallback option with OFF default before refresh"
 $null=Run 'M6Quiet';$null=Run 'FEMUiRefreshInputs';$null=Raw 'FEM24Install';$null=Raw 'FEM24Model'
 $front=$wb.Worksheets.Item('モデル');$xl.EnableEvents=$true;$front.Range('C4').Value2='メッシュ生成・支持条件'
 $optionRow=0;for($r=8;$r -le 30;$r++){if([string]$front.Cells($r,1).Value2 -like '*不適合Q8*'){$optionRow=$r;break}}
 Assert ($optionRow -gt 0) 'mesh fallback option visible in seven-screen model page'
 $front.Cells($optionRow,4).Value2=1;Assert ((Run 'M6ReadFallback') -eq '1') 'front ON writes canonical setting'
 $front.Cells($optionRow,4).Value2=0;Assert ((Run 'M6ReadFallback') -eq '0') 'front OFF writes canonical setting'
 $xl.EnableEvents=$false
 LoadFixture
 $n0=Snapshot '節点データ' 9;$e0=Snapshot '要素データ' 11
 $null=Run 'M6SetFallback' @(0);$null=Run 'M6ApplyFallback';$null=Run 'M6Validate' @(2,0,13)
 Assert ((Snapshot '節点データ' 9) -ceq $n0 -and (Snapshot '要素データ' 11) -ceq $e0) 'OFF leaves mesh unchanged'
 $conditions0=ConvertTo-Json -InputObject $wb.Worksheets.Item('節点データ').Range('D2:I14').Value2 -Compress
 $good=ConvertTo-Json -InputObject $wb.Worksheets.Item('要素データ').Range('A2:J2').Value2 -Compress
 $null=Run 'M6SetFallback' @(1);$null=Run 'M6PrimeQualityScale';$null=Run 'M6ApplyFallback';$null=Run 'M6Validate' @(1,2,14)
 Assert $true 'fallback mask uses current model size after a different model'
 Assert ((ConvertTo-Json -InputObject $wb.Worksheets.Item('要素データ').Range('A2:J2').Value2 -Compress) -ceq $good) 'ON preserves good Q8 connectivity and material'
 $originalConditions=ConvertTo-Json -InputObject $wb.Worksheets.Item('節点データ').Range('D2:I14').Value2 -Compress
 Assert ($originalConditions -ceq $conditions0) 'original nodal conditions preserved'
 $null=Run 'M6Preview';Assert ($wb.Worksheets.Item('図').Shapes.Count -gt 0) 'mixed mesh native figure with T6 legend'
 $solver=Run 'M6Run';Assert ($solver.StartsWith('PASS|')) 'mixed tapered patch solve, weight/reaction 45kN'
 $wb.Save();$wb.Close($false);$xl.EnableEvents=$true;$wb=$xl.Workbooks.Open($copy,0,$false);$xl.EnableEvents=$false
 $null=Run 'M6Validate' @(1,2,14);$null=Run 'M6Preview';Assert $true 'saved mixed topology reopens and renders'
 $null=Run 'M6ShuffledRows';Assert $true 'shuffled rows preserve saved type lookup and reject invalid input ordering'
 # Reject a different midpoint ID on the common edge, even at the same coordinates.
 $wn=$wb.Worksheets.Item('節点データ');$we=$wb.Worksheets.Item('要素データ');$mid=[int]$we.Cells(2,7).Value2
 $wn.Range('A16:I16').Value2=$wn.Range("A$($mid+1):I$($mid+1)").Value2;$wn.Range('A16').Value2=15
 $changed=$false;for($r=3;$r -le 4;$r++){foreach($c in 6,7,9){if([int]$we.Cells($r,$c).Value2 -eq $mid){$we.Cells($r,$c).Value2=15;$changed=$true;break}};if($changed){break}}
 $null=Raw 'FEMInvalidateRemeshedModel';$testError=[string](Raw 'M6Validate' @(1,2,15))
 Assert ($testError -like 'ERROR|*中間節点*不一致*') 'nonconforming midpoint ID rejected'
 LoadFixture;$wn=$wb.Worksheets.Item('節点データ');$wn.Range('B4').Value2=-0.5
 $beforeBadN=Snapshot '節点データ' 9;$beforeBadE=Snapshot '要素データ' 11
 $null=Run 'M6SetFallback' @(1);$testError=[string](Raw 'M6ApplyFallback')
 Assert ($testError.StartsWith('ERROR|')) 'invalid geometry rejected'
 Assert ((Snapshot '節点データ' 9) -ceq $beforeBadN -and (Snapshot '要素データ' 11) -ceq $beforeBadE) 'failed fallback leaves input tables unchanged'
 LoadFixture
 $null=Run 'M6Generate' @(0,$false);$offCount=$wb.Worksheets.Item('要素データ').Cells($wb.Worksheets.Item('要素データ').Rows.Count,1).End(-4162).Row-1
 $null=Run 'M6Validate' @($offCount,0,-1)
 $null=Run 'M6Generate' @(1,$false);$null=Run 'M6Validate' @(0,(2*$offCount),-1)
 Assert $true 'real shape-generation path repairs then converts quality failures'
 $null=Run 'M6Generate' @(1,$false,1);$null=Run 'M6Validate' @(-1,-1,-1)
 $we=$wb.Worksheets.Item('要素データ');$last=$we.Cells($we.Rows.Count,1).End(-4162).Row
 $t6Count=@($we.Range("K2:K$last").Value2 | Where-Object {$_ -eq 'T6'}).Count
 Assert ($t6Count -gt 0) 'Delaunay path converts quality failures after repair'

 $beforeBadN=Snapshot '節点データ' 9;$beforeBadE=Snapshot '要素データ' 11
 $testError=[string](Raw 'M6Generate' @(1,$true));Assert ($testError.StartsWith('ERROR|')) 'self-intersecting shape rejected'
 Assert ((Snapshot '節点データ' 9) -ceq $beforeBadN -and (Snapshot '要素データ' 11) -ceq $beforeBadE) 'invalid shape does not replace previous mesh'
 LoadFixture;$null=Run 'M6SetFallback' @(1);$null=Run 'M6ApplyFallback';$solver=Run 'M6Run';$null=Run 'M6Preview'
 $wb.Save()
 @{status='PASS';checks=@($checks);solver=$solver;book=$copy;source_sha256=$before;excel_version=$xl.Version;excel_build=$xl.Build}|ConvertTo-Json -Depth 8|Set-Content (Join-Path $Out 'results.json') -Encoding utf8
 if($Visible){
  Set-Content (Join-Path $Out 'ready.txt') 'Own mixed T6 test copy ready.'
  while(-not(Test-Path (Join-Path $Out 'close.txt'))){
   $command=Join-Path $Out 'command.txt'
   if(Test-Path $command){$name=(Get-Content $command -Raw).Trim();Remove-Item -LiteralPath $command
    switch($name){'model'{$null=Raw 'FEM24Model';$xl.EnableEvents=$true;$wb.Worksheets.Item('モデル').Range('C4').Value2='メッシュ生成・支持条件';$xl.EnableEvents=$false} 'mesh'{$null=Run 'M6Preview'} 'result'{$null=Raw 'FEMViewResult'}}
   }
   Start-Sleep -Milliseconds 300
  }
 }
 $wb.Close($false);$wb=$null
}catch{
 $_.ScriptStackTrace | Write-Output
 throw
}finally{
 if($wb){try{$wb.Close($false)}catch{}}
 if($xl){$xl.EnableEvents=$false;$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
if((Get-FileHash $Book).Hash -cne $before){throw 'Source workbook changed.'}
'ALL PASS'

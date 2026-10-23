param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Frame,[Parameter(Mandatory=$true)][string]$Branch,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process -Name EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; do not dispatch.'}
foreach($p in @($Book,$Frame,$Branch)){if(-not(Test-Path -LiteralPath $p)){throw "Path not found: $p"}}
$Book=(Resolve-Path $Book).Path;$Frame=(Resolve-Path $Frame).Path;$Branch=(Resolve-Path $Branch).Path
if(Test-Path -LiteralPath $Out){throw 'Preserve previous evidence.'}; New-Item -ItemType Directory $Out|Out-Null
$sourceHash=(Get-FileHash $Book).Hash
$testBook=Join-Path ([IO.Path]::GetFullPath($Out)) 'guard_input.xlsm'
Copy-Item -LiteralPath $Book -Destination $testBook
$dll=Join-Path (Split-Path $Book) 'soilfem_native_x64.dll'
if(Test-Path $dll){Copy-Item -LiteralPath $dll -Destination $Out}
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($testBook,0,$false)
 foreach($component in @($wb.VBProject.VBComponents)){if($component.Name -eq 'FSGuardDriver'){$wb.VBProject.VBComponents.Remove($component)}}
 $m=$wb.VBProject.VBComponents.Add(1);$m.Name='FSGuardDriver'
 $m.CodeModule.AddFromString(@'
Public Function FGLoad(ByVal folder As String) As String
 On Error GoTo Failed: SuppressUserMessages=True: P3LoadTimeFrame folder: FGLoad="PASS": Exit Function
Failed: FGLoad="ERROR|" & Err.Description
End Function
Public Function FGRun() As String
 On Error GoTo Failed: SuppressUserMessages=True: P0_RunAnalysis: FGRun=ResultStatus & "|" & AnalysisMessage: Exit Function
Failed: FGRun="ERROR|" & Err.Description
End Function
Public Function FGView() As String
 On Error GoTo Failed: FGView=CStr(FSViewPrefix()): Exit Function
Failed: FGView="ERROR|" & Err.Description
End Function
'@)
 $ctl=$xl.VBE.CommandBars.FindControl(1,578);if($ctl.Enabled){$ctl.Execute()}
 if($xl.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Guard driver compile failed'}
 $name=$wb.Name;$branchResult=''
 $null=$xl.Run("'$name'!FEMWriteTextSetting",'EXPORT_LOAD',$Branch);$branchResult=$xl.Run("'$name'!FGRun")
 if($branchResult -like 'PASS|*' -or ($branchResult -notmatch 'SRM照査枝は施工本線へ再開できません')){throw "Branch restart was not rejected: $branchResult"}
 $loaded=$xl.Run("'$name'!FGLoad",$Frame);if($loaded -ne 'PASS'){throw "Frame load failed: $loaded"}
 $mat=$wb.Worksheets.Item('材料データ');$oldMat=$mat.Range('B2').Value2;$mat.Range('B2').Value2=[double]$oldMat+1
 $materialResult=$xl.Run("'$name'!FGView");$mat.Range('B2').Value2=$oldMat
 if($materialResult -notlike 'ERROR|*' -or $materialResult -notmatch '一致|signature|入力|時刻別'){throw "Material signature guard missing: $materialResult"}
 $loaded=$xl.Run("'$name'!FGLoad",$Frame);if($loaded -ne 'PASS'){throw "Frame load failed: $loaded"}
 $water=$wb.Worksheets.Item('水条件');$oldAlpha=$water.Range('W12').Value2;$water.Range('W12').Value2=[double]$oldAlpha+0.001
 $alphaResult=$xl.Run("'$name'!FGView");$water.Range('W12').Value2=$oldAlpha
 if($alphaResult -notlike 'ERROR|*' -or $alphaResult -notmatch '一致|signature|入力|時刻別'){throw "Alpha signature guard missing: $alphaResult"}
 $loaded=$xl.Run("'$name'!FGLoad",$Frame);if($loaded -ne 'PASS'){throw "Frame load failed: $loaded"}
 $oldPorosity=$water.Range('AB12').Value2;$water.Range('AB12').Value2=[double]$oldPorosity+0.001
 $porosityResult=$xl.Run("'$name'!FGView");if($null -eq $oldPorosity){$water.Range('AB12').ClearContents()}else{$water.Range('AB12').Value2=$oldPorosity}
 if($porosityResult -notlike 'ERROR|*' -or $porosityResult -notmatch '一致|signature|入力|時刻別'){throw "Porosity signature guard missing: $porosityResult"}
 @{status='PASS';branch=$branchResult;material=$materialResult;alpha=$alphaResult;porosity=$porosityResult}|ConvertTo-Json|Set-Content (Join-Path $Out 'guards.json') -Encoding utf8
}catch{Write-Output $_.ScriptStackTrace;Write-Output $_.Exception.ToString();throw}finally{if($wb){try{$wb.Close($false)}catch{}};$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl);[GC]::Collect();[GC]::WaitForPendingFinalizers()}

if((Get-FileHash $Book).Hash -ne $sourceHash){throw 'Source evidence workbook changed'}

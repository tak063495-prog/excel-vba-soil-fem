param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Existing Excel; no test started.'}
$Book=(Resolve-Path $Book).Path;$Out=[IO.Path]::GetFullPath($Out)
if(Test-Path $Out){throw 'Preserve evidence.'}
New-Item -ItemType Directory -Path $Out | Out-Null
$copy=Join-Path $Out 'terrain_ui.xlsm';Copy-Item $Book $copy
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$true;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!";$xl.EnableEvents=$false
 $helper=$wb.VBProject.VBComponents.Add(1);$helper.Name='TerrainUiCheck'
 $helper.CodeModule.AddFromString(@'
Public Function TerrainUI(ByVal fail As Boolean) As String
 On Error GoTo Bad
 Dim ws As Worksheet, pts As Variant, i As Long
 SuppressUserMessages=True
 FEMWriteNumericSetting "MESH_GEN_COMPARE",2
 Set ws=ThisWorkbook.Worksheets("形状入力")
 ws.Range("B8:D2000").ClearContents:ws.Range("G8:J2000").ClearContents
 pts=Array(Array(1,0,0),Array(2,20,0),Array(3,20,4),Array(4,12,4),Array(5,6,8),Array(6,0,8))
 For i=0 To 5:ws.Cells(i+8,2).Value2=pts(i)(0):ws.Cells(i+8,3).Value2=pts(i)(1):ws.Cells(i+8,4).Value2=pts(i)(2):Next i
 ws.Cells(8,7).Value2=1:ws.Cells(8,8).Value2="外側":ws.Cells(8,9).Value2="1,2,3,4,5,6":ws.Cells(8,10).Value2=2
 If fail Then
  ws.Cells(11,3).Value2=6.4
  SuppressUserMessages=False
 End If
 FEMUiGenerateShape
 If fail Then
  TerrainUI=FEMUiMeshPreviewError
 Else
  If Not FEMShapeGenerationOK Or Len(FEMUiMeshPreviewError)>0 Or ActiveSheet.Name<>"図" Then Err.Raise vbObjectError+7000,,FEMUiMeshPreviewError
  TerrainUI="PASS|" & FEMReadTextSetting("MESH_GEN_LAST_METHOD","") & "|" & P5QualitySummary()
 End If
 Exit Function
Bad:TerrainUI="ERROR|" & Err.Description
End Function
'@)
 $ctl=$xl.VBE.CommandBars.FindControl(1,578);if($ctl.Enabled){$ctl.Execute()}
 if($xl.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Compile failed.'}
 $xl.Visible=$true;$xl.WindowState=-4143;$xl.Width=1100;$xl.Height=720
 $status=[string]$xl.Run($prefix+'TerrainUI',$false)
 if(-not $status.StartsWith('PASS|')){throw $status}
 $xl.ActiveWindow.Zoom=70
 @{status='READY';compile='PASS';result=$status;legend=[string]$wb.Worksheets.Item('メッシュ設定').Range('E29').Value2}|ConvertTo-Json|Set-Content (Join-Path $Out 'ui.json') -Encoding UTF8
 $settingsShown=$false
 while(-not(Test-Path (Join-Path $Out 'show_error.txt'))){
  if(-not $settingsShown -and (Test-Path (Join-Path $Out 'show_settings.txt'))){$wb.Worksheets.Item('メッシュ設定').Activate();$xl.ActiveWindow.Zoom=90;$xl.ActiveWindow.ScrollRow=28;$settingsShown=$true}
  Start-Sleep -Milliseconds 250
 }
 $err=[string]$xl.Run($prefix+'TerrainUI',$true)
 if($err.Length -lt 10 -or $err.StartsWith('PASS|')){throw 'Expected reason-specific error dialog.'}
 @{status='PASS';compile='PASS';result=$status;error=$err;legend=[string]$wb.Worksheets.Item('メッシュ設定').Range('E29').Value2}|ConvertTo-Json|Set-Content (Join-Path $Out 'ui.json') -Encoding UTF8
}finally{
 if($wb){$wb.Close($false)}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}

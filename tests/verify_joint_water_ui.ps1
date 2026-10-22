param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Existing Excel; do not dispatch.'}
$Book=(Resolve-Path $Book).Path;$Out=[IO.Path]::GetFullPath($Out)
if(Test-Path $Out){throw 'Preserve prior UI evidence.'}
New-Item -ItemType Directory -Path $Out | Out-Null
$copy=Join-Path $Out 'joint_water_ui.xlsm';Copy-Item $Book $copy
Copy-Item (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') $Out
$control=Join-Path $Out 'command.txt';$status=Join-Path $Out 'status.json'
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1
 $xl.EnableEvents=$true
 $wb=$xl.Workbooks.Open($copy,0,$false)
 $xl.EnableEvents=$false
 $module=$wb.VBProject.VBComponents.Add(1);$module.Name='JointWaterUiDriver'
 $module.CodeModule.AddFromString(@'
Public Function JWAction(ByVal action As String) As String
 On Error GoTo Failed
 SuppressUserMessages = True
 Select Case action
 Case "joint_input": FEMWaterShowJointInput
 Case "joint_result": FEMWaterShowJointResults
 Case "preview": SuppressUserMessages = False: FEMWaterPreview
 Case "result": FEMWaterResults
 Case "mesh": SuppressUserMessages = False: FEMUiPreviewMesh: If Len(FEMUiMeshPreviewError) > 0 Then Err.Raise vbObjectError + 1, , FEMUiMeshPreviewError
 Case "stress": FEMViewResult: If Len(FEMViewerLastError) > 0 Then Err.Raise vbObjectError + 1, , FEMViewerLastError
 End Select
 SuppressUserMessages = False
 JWAction = "PASS"
 Exit Function
Failed:
 SuppressUserMessages = False
 JWAction = "ERROR|" & Err.Description
End Function
'@)
 function Call([string]$macro){$xl.Run("'"+$wb.Name+"'!"+$macro)}
 function Snapshot([string]$file){
  $data=@{}
  foreach($n in '水条件','節点データ','要素データ','材料データ','ステージ','接合'){
   $s=$wb.Worksheets.Item($n);$last=if($n -eq '水条件'){415}else{[Math]::Max(2,$s.Cells($s.Rows.Count,1).End(-4162).Row)}
   $cols=switch($n){'水条件'{10};'節点データ'{9};'要素データ'{10};'材料データ'{13};'ステージ'{6};'接合'{7}}
   $v=$s.Range($s.Cells(1,1),$s.Cells($last,$cols)).Value2;$rows=[Collections.Generic.List[object]]::new()
   for($i=1;$i -le $last;$i++){$r=[Collections.Generic.List[object]]::new();for($j=1;$j -le $cols;$j++){$r.Add($v[$i,$j])};$rows.Add(@($r))}
   $data[$n]=@($rows)
  }
  $data|ConvertTo-Json -Depth 8|Set-Content (Join-Path $Out $file) -Encoding UTF8
 }
 Snapshot 'before_layout.json'
 $before=ConvertTo-Json -InputObject $wb.Worksheets.Item('接合').Range('A2:G2').Value2 -Compress
 $null=Call 'FEMWaterInstallUi';$null=Call 'FEMUiInvalidateLayout';$null=Call 'FEMUiUpdateLayout'
 if($before -ne (ConvertTo-Json -InputObject $wb.Worksheets.Item('接合').Range('A2:G2').Value2 -Compress)){throw 'Joint input changed during layout.'}
 Snapshot 'after_layout.json'
 $sig=Call 'GWInputSignature';$storedSig=$wb.Worksheets.Item('水圧結果').Range('R5').Value2
 if([string]$sig -ne [string]$storedSig){throw 'Saved result signature missing or changed during normal opening.'}
 $xl.Visible=$true;$xl.WindowState=-4143;$xl.Width=1050;$xl.Height=650
 $wb.Worksheets.Item('水条件').Activate();$xl.ActiveWindow.Zoom=80
 @{action='ready';status='PASS';book=$copy;input_preserved=$true;signature_now=$sig;signature_stored=$storedSig}|ConvertTo-Json|Set-Content $status -Encoding UTF8
 $previous=''
 while($true){
  $command=if(Test-Path $control){(Get-Content $control -Raw).Trim()}else{''}
  if($command -and $command -ne $previous){
   $previous=$command;$value='PASS'
   switch($command){
    'water_input'{$wb.Worksheets.Item('水条件').Activate();$xl.ActiveWindow.ScrollRow=1;$xl.ActiveWindow.ScrollColumn=1}
    'joint_input'{$value=$xl.Run("'"+$wb.Name+"'!JWAction",'joint_input');$xl.ActiveWindow.Zoom=80;$xl.Goto($wb.Worksheets.Item('接合').Range('F2'),$true);$xl.ActiveWindow.ScrollColumn=1}
    'joint_result'{$value=$xl.Run("'"+$wb.Name+"'!JWAction",'joint_result');$xl.ActiveWindow.Zoom=80}
    'preview'{$value=$xl.Run("'"+$wb.Name+"'!JWAction",'preview');$xl.ActiveWindow.Zoom=60;$xl.ActiveWindow.ScrollRow=1;$xl.ActiveWindow.ScrollColumn=1}
    'result'{$value=$xl.Run("'"+$wb.Name+"'!JWAction",'result');$xl.ActiveWindow.Zoom=60;$xl.ActiveWindow.ScrollRow=1;$xl.ActiveWindow.ScrollColumn=1}
    'reopen'{$wb.Save();$wb.Close($false);$xl.EnableEvents=$true;$wb=$xl.Workbooks.Open($copy,0,$false);$xl.EnableEvents=$false;$null=Call 'FEMWaterInstallUi';$value=$xl.Run("'"+$wb.Name+"'!JWAction",'joint_result')}
    'stale'{$xl.EnableEvents=$true;$wb.Worksheets.Item('接合').Range('F2').Value2='PRESCRIBED';$wb.Worksheets.Item('接合').Range('G2').Value2='W';$xl.EnableEvents=$false;$value=$xl.Run("'"+$wb.Name+"'!JWAction",'joint_result')}
    'mesh'{$value=$xl.Run("'"+$wb.Name+"'!JWAction",'mesh');$xl.ActiveWindow.Zoom=60}
    'stress'{$value=$xl.Run("'"+$wb.Name+"'!JWAction",'stress');$xl.ActiveWindow.Zoom=60}
    'quit'{}
    default{throw "Unknown action: $command"}
   }
   $record=@{action=$command;status=$value;sheet=$wb.ActiveSheet.Name;shapes=$wb.ActiveSheet.Shapes.Count;freeze_columns=$xl.ActiveWindow.SplitColumn;freeze_rows=$xl.ActiveWindow.SplitRow}
   $record|ConvertTo-Json|Set-Content $status -Encoding UTF8;$record|ConvertTo-Json -Compress|Add-Content (Join-Path $Out 'actions.jsonl') -Encoding UTF8
   if($command -eq 'quit'){break}
  }
  Start-Sleep -Milliseconds 250
 }
 $wb.Close($false)
}finally{
 foreach($reference in @($module,$wb)){if($reference -and [Runtime.InteropServices.Marshal]::IsComObject($reference)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reference)}catch{}}}
 $xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl);[GC]::Collect();[GC]::WaitForPendingFinalizers()
}

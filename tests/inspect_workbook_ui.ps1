param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$OutputDirectory)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no inspection instance started.'}
$src=(Resolve-Path -LiteralPath $Book).Path;$hash=(Get-FileHash -LiteralPath $src).Hash
if(Test-Path -LiteralPath $OutputDirectory){throw 'Use a new inspection directory.'}
$folder=(New-Item -ItemType Directory -Path $OutputDirectory).FullName
$copy=Join-Path $folder 'screen.xlsm';Copy-Item -LiteralPath $src -Destination $copy
$app=$null;$wb=$null
try{
 $app=New-Object -ComObject Excel.Application;$app.EnableEvents=$false;$app.DisplayAlerts=$false;$app.AutomationSecurity=1;$app.Visible=$true
 $app.VBE.MainWindow.Visible=$false
 $wb=$app.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!"
 $h=$wb.VBProject.VBComponents.Add(1);$h.Name='ScreenInspection'
 $h.CodeModule.AddFromString('Public Sub InspectPreview()'+"`r`n"+'SuppressUserMessages = True'+"`r`n"+'FEMUiPreviewMesh'+"`r`n"+'End Sub')
 $app.Run($prefix+'FEMWriteBuildStamp',$false)
 $app.Run($prefix+'InspectPreview');$wb.Save()
 $wb.Worksheets.Item('操作パネル').Activate();$app.ActiveWindow.Zoom=85;$app.ActiveWindow.ScrollRow=1
 Write-Output ('READY '+$folder+'; page.txt=sheet|scrollrow, finish to close this copy only')
 for($i=0;$i -lt 600;$i++){
  if(Test-Path (Join-Path $folder 'finish')){break}
  $request=Join-Path $folder 'page.txt'
  if(Test-Path -LiteralPath $request){
   $parts=(Get-Content -LiteralPath $request -Raw -Encoding UTF8).Trim().Split('|')
   $wb.Worksheets.Item($parts[0]).Activate();$app.ActiveWindow.Zoom=85;$app.ActiveWindow.ScrollRow=[int]$parts[1]
   Remove-Item -LiteralPath $request
  }
  Start-Sleep -Milliseconds 500
 }
}finally{
 if($wb){try{$wb.Close($false)}catch{}}
 if($app){try{$app.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)}
 [GC]::Collect();[GC]::WaitForPendingFinalizers()
}
if((Get-FileHash -LiteralPath $src).Hash -cne $hash){throw 'Source workbook changed.'}

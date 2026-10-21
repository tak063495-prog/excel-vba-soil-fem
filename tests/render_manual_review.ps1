param([string]$ResultBook='',[string]$Out='tests/tmp/manual_review/wall_render')
$ErrorActionPreference='Stop';if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running.'}
$repo=Split-Path -Parent $PSScriptRoot
$Out=[IO.Path]::GetFullPath($Out);New-Item -ItemType Directory -Force -Path $Out|Out-Null
if(-not $ResultBook){
 Add-Type -AssemblyName System.IO.Compression.FileSystem
 $archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $repo 'docs/manuals/construction-evidence/20261020/wall_davis08_workbook.zip'))
 try{$ResultBook=Join-Path $Out 'archived_source.xlsm';if(-not(Test-Path $ResultBook)){[IO.Compression.ZipFileExtensions]::ExtractToFile($archive.GetEntry('2DSoilFEM_ui_20261020.xlsm'),$ResultBook)}}finally{$archive.Dispose()}
}
$ResultBook=(Resolve-Path -LiteralPath $ResultBook).Path
$copy=Join-Path $Out 'render.xlsm';if(Test-Path -LiteralPath $copy){throw 'Render output already exists; choose a fresh Out directory.'};Copy-Item -LiteralPath $ResultBook -Destination $copy
$xl=$null;$wb=$null
try{
 $xl=New-Object -ComObject Excel.Application;$xl.EnableEvents=$false;$xl.DisplayAlerts=$false;$xl.AutomationSecurity=1
 $wb=$xl.Workbooks.Open($copy,0,$false);$prefix="'"+$wb.Name+"'!"
 foreach($m in 'FEMIncrementSvg','FEMStrainMath'){$cm=$wb.VBProject.VBComponents.Item($m).CodeModule;$cm.DeleteLines(1,$cm.CountOfLines);$cm.AddFromString([IO.File]::ReadAllText((Join-Path $repo "src/vba/$m.bas")))}
 # Test-only export hook: copy the exact SVG before the viewer deletes its temp file.
 $process=[Diagnostics.Process]::new();$process.StartInfo.FileName='git';$process.StartInfo.Arguments='show HEAD:src/vba/FEMViewer.bas';$process.StartInfo.WorkingDirectory=$repo;$process.StartInfo.UseShellExecute=$false;$process.StartInfo.RedirectStandardOutput=$true;$process.StartInfo.StandardOutputEncoding=[Text.Encoding]::UTF8
 [void]$process.Start();$viewer=$process.StandardOutput.ReadToEnd();$process.WaitForExit();if($process.ExitCode -ne 0){throw 'Cannot read tracked viewer source.'}
 $viewer=$viewer.Replace('If Len(tempPath) > 0 Then Kill tempPath','If Len(tempPath) > 0 Then FileCopy tempPath, ThisWorkbook.Path & "\latest.svg"'+"`r`n"+'    If Len(tempPath) > 0 Then Kill tempPath')
 $cm=$wb.VBProject.VBComponents.Item('FEMViewer').CodeModule;$cm.DeleteLines(1,$cm.CountOfLines);$cm.AddFromString($viewer)
 $xl.Run($prefix+'SetP0SilentMode',$true)|Out-Null
 foreach($mode in 'A','B'){
  $opt=$wb.Worksheets.Item('出力・表示');$opt.Unprotect();$opt.Range('C27').Value2=if($mode -eq 'A'){'A：選択結果へ等値線を追加'}else{'B：増分専用カラー'}
  $opt.Range('J28').Value2=$true;$opt.Range('J29').Value2=$true
  $xl.Run($prefix+'FEMViewResult')|Out-Null
  if(-not(Test-Path (Join-Path $Out 'latest.svg'))){throw 'No SVG generated from saved results.'}
  Move-Item -LiteralPath (Join-Path $Out 'latest.svg') -Destination (Join-Path $Out "actual_$mode.svg") -Force
 }
 Write-Output 'PASS A/B redrawn by Excel from unchanged archived result tables; no solve'
}finally{if($wb){$wb.Close($false)};if($xl){$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)};[GC]::Collect();[GC]::WaitForPendingFinalizers()}

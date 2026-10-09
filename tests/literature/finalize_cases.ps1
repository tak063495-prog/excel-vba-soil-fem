param([Parameter(Mandatory=$true)][string]$OutputRoot)
$ErrorActionPreference='Stop'
$visualTask=Join-Path $OutputRoot 'visual_checks'
New-Item -ItemType Directory -Force -Path $visualTask|Out-Null
$xlTask=New-Object -ComObject Excel.Application
$xlTask.EnableEvents=$false;$xlTask.DisplayAlerts=$false;$xlTask.AutomationSecurity=1
$checksTask=[Collections.Generic.List[object]]::new()
$harnessTask=@'
Public Function LiteratureViewState() As Variant
  LiteratureViewState = Array(FEMViewerElementCount, FEMViewerLastError)
End Function
'@
try{
  foreach($fileTask in Get-ChildItem -LiteralPath (Join-Path $OutputRoot 'results') -Filter '*.json' -File){
    $resultTask=Get-Content -LiteralPath $fileTask.FullName -Raw -Encoding UTF8|ConvertFrom-Json
    if($null -eq $resultTask.metrics){continue}
    $pathTask=Join-Path $OutputRoot ('cases/'+$resultTask.case+'.xlsm')
    $wbTask=$xlTask.Workbooks.Open((Resolve-Path -LiteralPath $pathTask).Path,0,$false)
    try{
      $moduleTask=$wbTask.VBProject.VBComponents.Add(1);$moduleTask.Name='LiteratureViewHarness';$moduleTask.CodeModule.AddFromString($harnessTask)
      $prefixTask="'"+$wbTask.Name+"'!"
      $xlTask.Run($prefixTask+'SetP0SilentMode',$true)
      $sourceUrlTask=$(if($resultTask.reference.primary_url){[string]$resultTask.reference.primary_url}elseif($resultTask.reference.source_id -like 'griffiths_lane_1999*'){'https://inside.mines.edu/~vgriffit/slope64/Griffiths%20and%20Lane%201999'}elseif($resultTask.reference.source_id -eq 'pruska_homogeneous'){'https://data.fine.cz/handbooks-chapter-pdf/16_comparison_of_geotechnic_softwares_geo_fem_plaxis_z-soil.pdf'}else{'https://docs.itascacg.com/itasca900/common/models/elastic/doc/modelelastic.html'})
      $materialTask=$wbTask.Worksheets.Item('材料データ')
      $materialTask.Range('O11').Value2='参照資料'
      $materialTask.Range('P11').Value2=$sourceUrlTask
      $materialTask.Columns.Item('O').ColumnWidth=14
      $materialTask.Columns.Item('P').ColumnWidth=72
      $materialTask.Range('P9:P11').Font.Size=9
      $materialTask.Range('P11').WrapText=$true
      $materialTask.Range('P11').Font.Size=9
      $materialTask.Rows.Item(11).RowHeight=60
      [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($materialTask)
      if($resultTask.reference.source_id -eq 'elastic_validation'){
        $selectionTask=$(if($resultTask.case -eq 'simple_shear'){3}else{2})
        $xlTask.Run($prefixTask+'FEMWriteNumericSetting','VIEW_COLOR_RESULT',$selectionTask)
      }
      foreach($modeTask in @('Initial','Result')){
        $xlTask.Run($prefixTask+('FEMView'+$modeTask))
        $stateTask=@($xlTask.Run($prefixTask+'LiteratureViewState'))
        $expectedTask=[int]$resultTask.metrics[16]
        if($modeTask -eq 'Result' -and -not $resultTask.status.StartsWith('PASS|')){
          $checksTask.Add(@{case=$resultTask.case;mode='ResultRejected';elements=$stateTask[0];error=$stateTask[1];ok=($stateTask[0] -eq 0 -and [string]$stateTask[1] -ne '')})
        }else{
          $checksTask.Add(@{case=$resultTask.case;mode=$modeTask;elements=$stateTask[0];error=$stateTask[1];ok=($stateTask[0] -eq $expectedTask -and [string]$stateTask[1] -eq '')})
        }
        if(($modeTask -eq 'Initial' -and $resultTask.case -like '*_critical*') -or ($modeTask -eq 'Result' -and $resultTask.status.StartsWith('PASS|') -and ($resultTask.case -like '*_critical*' -or $resultTask.case -in @('confined_combined','griffiths_1999_coarse_davis','pruska_h7_phi10_davis','pruska_h10.5_phi10_davis')))){
          $wsTask=$wbTask.Worksheets.Item('図');$setupTask=$wsTask.PageSetup
          $setupTask.PrintArea='A2:W45';$setupTask.Orientation=2;$setupTask.Zoom=$false;$setupTask.FitToPagesWide=1;$setupTask.FitToPagesTall=1
          $pdfTask=Join-Path $visualTask ($resultTask.case+'_'+$modeTask.ToLowerInvariant()+'.pdf')
          $wsTask.ExportAsFixedFormat(0,[IO.Path]::GetFullPath($pdfTask))
          [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($setupTask)
          [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wsTask)
          $wsTask=$wbTask.Worksheets.Item('材料データ');$setupTask=$wsTask.PageSetup
          $setupTask.PrintArea='A1:P12';$setupTask.Orientation=2;$setupTask.Zoom=$false;$setupTask.FitToPagesWide=1;$setupTask.FitToPagesTall=1
          $pdfTask=Join-Path $visualTask ($resultTask.case+'_inputs.pdf')
          $wsTask.ExportAsFixedFormat(0,[IO.Path]::GetFullPath($pdfTask))
          [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($setupTask)
          [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wsTask)
        }
      }
      # Save the initial mesh view, so opening a case shows its actual geometry.
      $xlTask.Run($prefixTask+'FEMViewInitial')
      $wbTask.VBProject.VBComponents.Remove($moduleTask)
      [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($moduleTask)
      $wbTask.Save()
    }finally{$wbTask.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wbTask)}
  }
}finally{$xlTask.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask);[GC]::Collect();[GC]::WaitForPendingFinalizers()}
$failTask=@($checksTask|Where-Object {-not $_.ok})
@{ok=($failTask.Count -eq 0);checks=$checksTask}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $visualTask 'viewer_checks.json') -Encoding UTF8
Write-Host ('Viewer checks='+$checksTask.Count+' failures='+$failTask.Count)
if($failTask.Count -gt 0){$failTask|ConvertTo-Json;exit 1}

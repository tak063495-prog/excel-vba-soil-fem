$ErrorActionPreference='Stop'
$rootTask=Split-Path -Parent $PSScriptRoot
Push-Location $rootTask
try {
$resultsTask=[Collections.Generic.List[object]]::new()
function PutRows($wb,$name,$rows,$cols){
  $ws=$wb.Worksheets.Item($name)
  $ws.Range($ws.Cells(2,1),$ws.Cells(2000,$cols)).ClearContents()
  $data=[object[,]]::new($rows.Count,$cols)
  for($r=0;$r -lt $rows.Count;$r++){for($c=0;$c -lt $cols;$c++){$data[$r,$c]=$rows[$r][$c]}}
  $ws.Range($ws.Cells(2,1),$ws.Cells($rows.Count+1,$cols)).Value2=$data
}
function Setting($wb,$key,$value){
  $method=$(if($value -is [string]){'FEMWriteTextSetting'}else{'FEMWriteNumericSetting'})
  $xlTask.Run("'"+$wb.Name+"'!"+$method,$key,$value)
}
$harnessTask=@'
Public Function FeatureRun() As String
  On Error GoTo Failed
  P0_RunAnalysis
  FeatureRun = ResultStatus & "|" & CStr(AnalysisOK) & "|" & CStr(RelativeResidualFree) & "|" & CStr(P3SuccessfulIncrementCount) & "|" & CStr(P6FactorizationCount) & "|" & CStr(P6FactorizationReuseCount) & "|" & CStr(NumberOfNode) & "|" & CStr(NumberOfElement) & "|" & AnalysisMessage
  Exit Function
Failed:
  FeatureRun = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Function FeatureMesh(ByVal shapeMode As Boolean) As String
  On Error GoTo Failed
  SetP0SilentMode True
  If shapeMode Then
    Call ボタン19_Click
  Else
    Call ボタン1_Click
  End If
  FeatureMesh = "PASS"
  Exit Function
Failed:
  FeatureMesh = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
Public Function FeatureStressMax() As Double
  Dim e As Long, gp As Long, i As Long, v As Double
  For e = 0 To NumberOfElement - 1
    For gp = 0 To 3
      For i = 0 To 2
        v = Abs(Elem(e).Stmat(i, gp))
        If v > FeatureStressMax Then FeatureStressMax = v
      Next i
    Next gp
  Next e
End Function
Public Function FeatureYoung() As Double
  FeatureYoung = Material(0).Young
End Function
Public Function FeatureDisp() As Variant
  FeatureDisp = TDisp
End Function
Public Function FeatureViewer(ByVal viewMode As String) As String
  On Error GoTo Failed
  SetP0SilentMode True
  Select Case viewMode
    Case "RESULT": FEMViewResult
    Case "ALL": FEMViewAll
    Case "INITIAL": FEMViewInitial
  End Select
  If Len(FEMViewerLastError) > 0 Then Err.Raise vbObjectError + 3950, , FEMViewerLastError
  FeatureViewer = "PASS|" & CStr(FEMViewerElementCount)
  Exit Function
Failed:
  FeatureViewer = "ERROR|" & Err.Description
End Function
Public Function FeatureShape() As String
  On Error GoTo Failed
  Application.Run "'" & ThisWorkbook.Name & "'!FEMUiBuildGeometry"
  FeatureShape = "PASS"
  Exit Function
Failed:
  FeatureShape = "ERROR|" & CStr(Err.Number) & "|" & Err.Description
End Function
'@
$xlTask=New-Object -ComObject Excel.Application
try {
  $xlTask.EnableEvents=$false;$xlTask.DisplayAlerts=$false;$xlTask.AutomationSecurity=1
  foreach($editionTask in @('after')){
    $srcTask=$(if($editionTask -eq 'before'){'2DSoilFEM_20261008_adaptive_policy5.xlsm'}else{'workbook/2DSoilFEM_20261008_practical.xlsm'})
    $copyTask=Join-Path $pwd ('tests/tmp/practical_features_'+$editionTask+'.xlsm')
    Copy-Item -LiteralPath $srcTask -Destination $copyTask -Force
    $wbTask=$xlTask.Workbooks.Open($copyTask,0,$false)
    $tmTask=$wbTask.VBProject.VBComponents.Add(1);$tmTask.Name='FeatureHarness';$tmTask.CodeModule.AddFromString($harnessTask)
    foreach($variantTask in @('joint_birth_death_skip','joint_birth_death_include','load_unload','matset','reset_u','reset_stress','disp_soil','disp_unload','disp_rcm')){
      foreach($keyTask in @('ACCEL_V1_PREDICTOR','ACCEL_V2A_REUSE','ACCEL_V2B_COST','ACCEL_V3_COST_SEARCH','ACCEL_V4_ANDERSON','ACCEL_V5_GMRES_LU','ACCEL_STEP_RECOVERY','ACCEL_ADAPTIVE')){Setting $wbTask $keyTask 0}
      Setting $wbTask 'EXPORT_MODE' 'OFF'; Setting $wbTask 'EXPORT_LOAD' '';Setting $wbTask 'DEBUG_MODE' 'OFF'
      Setting $wbTask 'SRM_FIXED_FS' 0;Setting $wbTask 'RCM_POLICY' 'OFF'
      Setting $wbTask 'MESH_BC_BOTTOM' 'PINNED';Setting $wbTask 'MESH_BC_LEFT' 'ROLLER';Setting $wbTask 'MESH_BC_RIGHT' 'NONE';Setting $wbTask 'MESH_BC_TOP' 'NONE';Setting $wbTask 'MESH_BC_PIN_CORNER' 'NONE'
      Setting $wbTask 'MESH_NX' 2;Setting $wbTask 'MESH_NY' 1;Setting $wbTask 'MESH_MATERIAL' 1
      foreach($pTask in @(@(1,0,0),@(2,0,1),@(3,1,1),@(4,1,0))){Setting $wbTask ('MESH_POINT'+$pTask[0]+'_X') ([double]$pTask[1]);Setting $wbTask ('MESH_POINT'+$pTask[0]+'_Y') ([double]$pTask[2])}
      $wbTask.Worksheets.Item('接合').Range('A2:E2000').ClearContents()
      $wbTask.Worksheets.Item('載荷').Range('A2:J2000').ClearContents()
      $matRowsTask=@(,@(1.0,14000.0,0.3,1.0,20.0,30.0,100.0,0.0,1.0,'ON','SOIL',$null,$null))
      if($variantTask -eq 'birth_death'){
        $matRowsTask+=,@(2.0,14000.0,0.3,1.0,20.0,30.0,100.0,0.0,1.0,'OFF','SOIL',$null,$null)
      } elseif($variantTask -like 'joint*'){
        $matRowsTask[0][10]='STRUCT'
        $matRowsTask+=,@(2.0,14000.0,0.3,1.0,20.0,30.0,100.0,0.0,1.0,'ON','SOIL',$null,$null)
        $matRowsTask+=,@(3.0,100000.0,0.3,1.0,0.0,0.0,10000.0,0.0,0.0,'ON','JOINT',100000.0,1.0)
      }
      if($variantTask -like 'joint_birth*'){
        $matRowsTask[2][9]='OFF'
        Setting $wbTask 'OUTPUT_STAGE_MODE' 'ALL'
        Setting $wbTask 'OUTPUT_INACTIVE_ELEMENTS' $(if($variantTask.EndsWith('include')){'INCLUDE'}else{'SKIP'})
      } else {
        $matRowsTask[0][4]=[double]0
        if($variantTask -like 'reset*'){$matRowsTask[0][10]='STRUCT'}
        Setting $wbTask 'OUTPUT_STAGE_MODE' 'ALL'
        Setting $wbTask 'OUTPUT_INACTIVE_ELEMENTS' 'SKIP'
      }
      PutRows $wbTask '材料データ' $matRowsTask 13
      if($variantTask -eq 'automesh'){
        Setting $wbTask 'MESH_GEN_COMPARE' 2
        if($editionTask -eq 'after'){
          $wsTask=$wbTask.Worksheets.Item('形状入力');$wsTask.Range('B8:D2000').ClearContents();$wsTask.Range('G8:J2000').ClearContents()
          $coordsTask=[object[,]]::new(4,3)
          $ptsTask=@(@(1.0,0.0,0.0),@(2.0,0.0,1.0),@(3.0,1.0,1.0),@(4.0,1.0,0.0))
          for($rTask=0;$rTask -lt 4;$rTask++){for($cTask=0;$cTask -lt 3;$cTask++){$coordsTask[$rTask,$cTask]=$ptsTask[$rTask][$cTask]}}
          $wsTask.Range('B8:D11').Value2=$coordsTask
          $wsTask.Range('G8').Value2=[double]1;$wsTask.Range('H8').Value2='外側';$wsTask.Range('I8').Value2='1,2,3,4';$wsTask.Range('J8').Value2=[double]0.5
          $shapeResultTask=$xlTask.Run("'"+$wbTask.Name+"'!FeatureShape");if($shapeResultTask -ne 'PASS'){throw $shapeResultTask}
        } else {
          $wsTask=$wbTask.Worksheets.Item('図形定義');$wsTask.Range('A1:N2000').ClearContents()
          $ptsTask=@(@(1.0,0.0,0.0),@(2.0,0.0,1.0),@(3.0,1.0,1.0),@(4.0,1.0,0.0))
          for($rTask=0;$rTask -lt 4;$rTask++){for($cTask=0;$cTask -lt 3;$cTask++){$wsTask.Cells($rTask+2,$cTask+1).Value2=$ptsTask[$rTask][$cTask]}}
          $wsTask.Range('H1').Value2=[double]4;$wsTask.Range('J1').Value2=[double]0;$wsTask.Range('L1').Value2=[double]4;$wsTask.Range('N1').Value2=[double]0.5
          $wsTask.Range('D3').Value2=[double]1;$wsTask.Range('D24').Value2=[double]0;$wsTask.Range('E3').Value2=[double]4;$wsTask.Range('F3').Value2=[double]0.5
          for($cTask=1;$cTask -le 4;$cTask++){$wsTask.Cells(3,$cTask+6).Value2=[double]$cTask}
        }
      }
      $meshResultTask=$xlTask.Run("'"+$wbTask.Name+"'!FeatureMesh",($variantTask -eq 'automesh'));if($meshResultTask -ne 'PASS'){throw $meshResultTask}
      if(($variantTask -eq 'birth_death' -or $variantTask -like 'joint*')){$wbTask.Worksheets.Item('要素データ').Cells(3,10).Value2=[double]2}
      if($variantTask -like 'joint*'){PutRows $wbTask '接合' (,@(1.0,2.0,3.0,1.0,'共有辺')) 5}
      $stageRowsTask=@(,@(1.0,'GRAVITY',$null,$null,1.0,'自重'))
      if($variantTask -eq 'birth_death'){
        $stageRowsTask+=,@(2.0,'BIRTH',2.0,$null,1.0,'追加')
        $stageRowsTask+=,@(3.0,'GRAVITY',$null,$null,1.0,'追加後')
        $stageRowsTask+=,@(4.0,'DEATH',2.0,$null,1.0,'撤去')
        $stageRowsTask+=,@(5.0,'GRAVITY',$null,$null,1.0,'撤去後')
      }
      if($variantTask -like 'joint_birth*'){
        $stageRowsTask+=,@(2.0,'BIRTH',3.0,$null,1.0,'接合を追加')
        $stageRowsTask+=,@(3.0,'GRAVITY',$null,$null,1.0,'接合後')
        $stageRowsTask+=,@(4.0,'DEATH',3.0,$null,1.0,'接合を解除')
        $stageRowsTask+=,@(5.0,'GRAVITY',$null,$null,1.0,'解除後')
      } elseif($variantTask -eq 'load_unload'){
        PutRows $wbTask '載荷' (,@(1.0,'TOP','Y','FORCE',-1.0,$null,$null,$null,$null,'上面荷重')) 10
        $stageRowsTask+=,@(2.0,'LOAD',1.0,$null,1.0,'載荷')
        $stageRowsTask+=,@(3.0,'UNLOAD',1.0,$null,1.0,'除荷')
      } elseif($variantTask -eq 'matset'){
        $stageRowsTask+=,@(2.0,'MATSET',1.0,'E=28000',1.0,'材料変更')
        $stageRowsTask+=,@(3.0,'GRAVITY',$null,$null,1.0,'変更後')
      } elseif($variantTask -like 'reset*' -or $variantTask -like 'disp*'){
        PutRows $wbTask '載荷' (,@(1.0,'TOP','Y','DISP',-0.001,$null,$null,$null,$null,'上面変位')) 10
        $stageRowsTask=@(,@(1.0,'LOAD',1.0,$null,1.0,'変位載荷'))
        if($variantTask -like 'reset*'){$stageRowsTask+=,@(2.0,$variantTask.ToUpper(),$null,$null,1.0,'リセット')}
        if($variantTask -eq 'disp_unload'){$stageRowsTask+=,@(2.0,'UNLOAD',1.0,$null,1.0,'変位除荷')}
        if($variantTask -eq 'disp_rcm'){Setting $wbTask 'RCM_POLICY' 'ON'}
      }
      PutRows $wbTask 'ステージ' $stageRowsTask 6
      $runTask=$xlTask.Run("'"+$wbTask.Name+"'!FeatureRun")
      $dispTask=@($xlTask.Run("'"+$wbTask.Name+"'!FeatureDisp"))
      Write-Output ($editionTask+' '+$variantTask+' '+$runTask)
      if(-not $runTask.StartsWith('PASS|True') -and ($variantTask -ne 'joint' -or $editionTask -eq 'after')){throw $runTask}
      if($variantTask -like 'joint_birth*'){
        $outTask=$wbTask.Worksheets.Item('接合結果')
        $countTask=$outTask.Cells($outTask.Rows.Count,1).End(-4162).Row-1
        $expectedTask=$(if($variantTask.EndsWith('include')){15}else{6})
        if($countTask -ne $expectedTask){throw ('joint stage row count '+$countTask)}
        for($rowTask=2;$rowTask -le $countTask+1;$rowTask++){
          $stageTask=[int]$outTask.Cells($rowTask,1).Value2
          $isActiveTask=$stageTask -in @(2,3)
          if($outTask.Cells($rowTask,13).Value2 -ne [int]$isActiveTask){throw 'wrong joint stage active flag'}
          if(-not $isActiveTask -and $null -ne $outTask.Cells($rowTask,7).Value2){throw 'inactive joint traction present'}
        }
      } elseif($variantTask -in @('load_unload','reset_u','disp_unload')){
        $maxTask=($dispTask|ForEach-Object {[Math]::Abs($_)}|Measure-Object -Maximum).Maximum
        if($maxTask -gt 1e-9){throw ('displacement did not return to zero '+$maxTask)}
      } elseif($variantTask -eq 'reset_stress'){
        if($xlTask.Run("'"+$wbTask.Name+"'!FeatureStressMax") -gt 1e-9){throw 'stress reset did not clear stress'}
        $savedNodesTask=$wbTask.Worksheets.Item('結果節点')
        $lastNodeRowTask=$savedNodesTask.Cells($savedNodesTask.Rows.Count,1).End(-4162).Row
        $referenceYTask=@()
        for($nrTask=2;$nrTask -le $lastNodeRowTask;$nrTask++){
          if($savedNodesTask.Cells($nrTask,1).Value2 -eq 2){$referenceYTask+= [double]$savedNodesTask.Cells($nrTask,4).Value2}
        }
        if([Math]::Abs(($referenceYTask|Measure-Object -Maximum).Maximum - 0.999) -gt 1e-10){throw 'RESET_STRESS reference geometry incorrect'}
        $wbTask.Save();$wbTask.Close($false)
        $wbTask=$xlTask.Workbooks.Open($copyTask,0,$false)
      } elseif($variantTask -eq 'matset'){
        if($xlTask.Run("'"+$wbTask.Name+"'!FeatureYoung") -ne 28000){throw 'MATSET did not update E'}
      } elseif($variantTask -like 'disp*'){
        $maxTask=($dispTask|ForEach-Object {[Math]::Abs($_)}|Measure-Object -Maximum).Maximum
        if([Math]::Abs($maxTask - 0.001) -gt 1e-10){throw 'prescribed displacement magnitude incorrect'}
        if($xlTask.Run("'"+$wbTask.Name+"'!FeatureStressMax") -lt 1){throw 'prescribed displacement produced no stress'}
      }
      $viewsTask=@()
      if($editionTask -eq 'after'){
        foreach($modeTask in @('RESULT','ALL','INITIAL')){
          $vTask=$xlTask.Run("'"+$wbTask.Name+"'!FeatureViewer",$modeTask)
          if(-not $vTask.StartsWith('PASS|')){throw $vTask}
          $viewsTask+=@{Mode=$modeTask;Result=$vTask}
        }
      }
      if($variantTask -eq 'reset_stress'){
        $xlTask.EnableEvents=$true
        $wbTask.Worksheets.Item('節点データ').Cells(2,2).Value2=0.25
        $xlTask.EnableEvents=$false
        $dirtyViewTask=$xlTask.Run("'"+$wbTask.Name+"'!FeatureViewer",'RESULT')
        if(-not $dirtyViewTask.Contains('入力変更後')){throw ('RESET_STRESS dirty input safeguard failed '+$dirtyViewTask)}
      }
      $resultsTask.Add([PSCustomObject]@{Edition=$editionTask;Variant=$variantTask;Metrics=$runTask;Disp=$dispTask;Views=$viewsTask})
    }
    $wbTask.Close($false)
  }
  $resultsTask | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath tests/results/practical_workflow_checks_20261008.json -Encoding UTF8
} finally {
  if($null -ne $wbTask){try{$wbTask.Close($false)}catch{}}
  try{$xlTask.Quit()}catch{}
  [Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)|Out-Null
}

} finally { Pop-Location }

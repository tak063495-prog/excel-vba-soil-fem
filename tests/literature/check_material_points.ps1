param([Parameter(Mandatory=$true)][string]$SourceWorkbook,[Parameter(Mandatory=$true)][string]$OutputRoot,[switch]$RepairSelfTestInputs)
$ErrorActionPreference='Stop'
$sourceHashTask=(Get-FileHash -LiteralPath $SourceWorkbook -Algorithm SHA256).Hash
New-Item -ItemType Directory -Force -Path $OutputRoot|Out-Null
$pathTask=Join-Path $OutputRoot 'material_point_scratch.xlsm'
Copy-Item -LiteralPath $SourceWorkbook -Destination $pathTask -Force
$codeTask=@'
Public Function LiteratureSelfRows() As Variant
  Dim ws As Worksheet, a() As Variant, r As Long, c As Long, k As Long
  Set ws = GetIntegratedResultSheet()
  ReDim a(0 To 599)
  For r = 1 To 50
    For c = 40 To 51
      a(k) = ws.Cells(r, c).Value2: k = k + 1
    Next c
  Next r
  LiteratureSelfRows = a
End Function
Public Function LiteraturePoint(ByVal phi As Double, ByVal ey As Double, ByVal gxy As Double) As Variant
  LiteraturePoint = LiteratureProbe(phi, 0#, ey, gxy, 0#, 0#, 0#, 0#)
End Function
Public Function LiteratureProbe(ByVal phi As Double, ByVal ex As Double, ByVal ey As Double, ByVal gxy As Double, ByVal sx0 As Double, ByVal sy0 As Double, ByVal tau0 As Double, ByVal sz0 As Double) As Variant
  Dim ip As P2_MaterialPointInput, op As P2_MaterialPointOutput, ok As Boolean
  ip.Young = 100000#: ip.Poisson = 0.3: ip.cohesion = 10#
  ip.frictionAngle = phi: ip.dilationAngle = 0#
  ip.StrainIncrement(0) = ex: ip.StrainIncrement(1) = ey: ip.StrainIncrement(2) = gxy
  ip.PreviousStress(0) = sx0: ip.PreviousStress(1) = sy0: ip.PreviousStress(2) = tau0: ip.PreviousStress(3) = sz0
  ip.tolerance = 0.000000001: ip.maxIterations = 50: ip.maxSubsteps = 1
  ip.EnableSubstepping = False
  ok = P2MaterialPointUpdate(ip, op)
  LiteratureProbe = Array(ok, op.converged, op.Elastic, op.failureCode, op.failureMessage, op.Stress(0), op.Stress(1), op.Stress(2), op.Stress(3), op.PlasticStrain(0), op.PlasticStrain(1), op.PlasticStrain(2), op.PlasticStrain(3), op.YieldFunction, op.PlasticMultiplier, op.Tangent(0, 0), op.Tangent(0, 1), op.Tangent(0, 2), op.Tangent(1, 0), op.Tangent(1, 1), op.Tangent(1, 2), op.Tangent(2, 0), op.Tangent(2, 1), op.Tangent(2, 2))
End Function
'@
$xlTask=New-Object -ComObject Excel.Application
$xlTask.EnableEvents=$false;$xlTask.DisplayAlerts=$false;$xlTask.AutomationSecurity=1
$wbTask=$null;$recordsTask=[Collections.Generic.List[object]]::new()
function ProbeTask($phi,$strain,$previous){
  return @($xlTask.Run($prefixTask+'LiteratureProbe',[double]$phi,[double]$strain[0],[double]$strain[1],[double]$strain[2],[double]$previous[0],[double]$previous[1],[double]$previous[2],[double]$previous[3]))
}
try{
  $wbTask=$xlTask.Workbooks.Open([IO.Path]::GetFullPath($pathTask),0,$false)
  if($RepairSelfTestInputs){
    # Test-only trial patch in this disposable copy. The material update and
    # solver procedures stay unchanged, and the workbook is never saved.
    $engineTask=$wbTask.VBProject.VBComponents.Item('FEMEngine').CodeModule
    foreach($oldTask in @('  P2SetSelfTestInput inputState, -0.1, 0#, 0#, 30#, 1#, 0#','  P2SetSelfTestInput inputState, -0.05, 0#, 0#, 30#, 0#, 0#')){
      $allTask=$engineTask.Lines(1,$engineTask.CountOfLines) -split "`r`n"
      $indicesTask=@(0..($allTask.Count-1)|Where-Object {$allTask[$_] -ceq $oldTask})
      if($indicesTask.Count -ne 1){throw 'Self-test input patch expected one exact match'}
      $newTask=$oldTask.Replace('-0.1, 0#, 0#','0.01, -0.009, 0#').Replace('-0.05, 0#, 0#','0.01, -0.009, 0#')
      $engineTask.ReplaceLine($indicesTask[0]+1,$newTask)
    }
  }
  $componentTask=$wbTask.VBProject.VBComponents.Add(1);$componentTask.Name='LiteraturePointHarness';$componentTask.CodeModule.AddFromString($codeTask)
  $prefixTask="'"+$wbTask.Name+"'!"
  $xlTask.Run($prefixTask+'SetP0SilentMode',$true)
  $selfTask=$xlTask.Run($prefixTask+'P2RunMaterialPointSelfTest')
  $selfRowsTask=@($xlTask.Run($prefixTask+'LiteratureSelfRows'))
  $edgeTask=$xlTask.Run($prefixTask+'P2RunEdgeReturnSelfTest')
  foreach($rowTask in @(@('confined_phi0_elastic',0,0.0001,0),@('confined_phi0_plastic',0,0.01,0),@('confined_phi20_elastic',20,0.0001,0),@('confined_phi20_plastic',20,0.003,0),@('confined_phi20_high_plastic',20,0.01,0),@('pure_shear_elastic',0,0,0.0001),@('pure_shear_plastic',0,0,0.001))){
    $valuesTask=@($xlTask.Run($prefixTask+'LiteraturePoint',[double]$rowTask[1],[double]$rowTask[2],[double]$rowTask[3]))
    $probesTask=[Collections.Generic.List[object]]::new()
    $pathsTask=@(@{name='total_increment';strain=@(0,[double]$rowTask[2],[double]$rowTask[3]);previous=@(0,0,0,0);h=1e-8})
    if(-not [bool]$valuesTask[2]){
      $tinyTask=$(if([double]$rowTask[3] -ne 0){@(0,0,1e-7)}else{@(0,1e-7,0)})
      $pathsTask+=@{name='small_reload';strain=$tinyTask;previous=@($valuesTask[5],$valuesTask[6],$valuesTask[7],$valuesTask[8]);h=1e-9}
      $pathsTask+=@{name='small_reload_h3';strain=$tinyTask;previous=@($valuesTask[5],$valuesTask[6],$valuesTask[7],$valuesTask[8]);h=3e-9}
    }
    foreach($probeTask in $pathsTask){
      $baseTask=ProbeTask $rowTask[1] $probeTask.strain $probeTask.previous
      $fdTask=[Collections.Generic.List[object]]::new();$branchTask=[Collections.Generic.List[object]]::new();$okTask=[bool]$baseTask[0]
      for($colTask=0;$colTask -lt 3;$colTask++){
        $plusTask=@($probeTask.strain);$minusTask=@($probeTask.strain)
        $plusTask[$colTask]+=$probeTask.h;$minusTask[$colTask]-=$probeTask.h
        $vpTask=ProbeTask $rowTask[1] $plusTask $probeTask.previous
        $vmTask=ProbeTask $rowTask[1] $minusTask $probeTask.previous
        $branchTask.Add(@{column=$colTask;plus_values=$vpTask;minus_values=$vmTask})
        $okTask=$okTask -and [bool]$vpTask[0] -and [bool]$vmTask[0]
        $derivativeTask=[double[]]::new(3)
        for($rTask=0;$rTask -lt 3;$rTask++){$derivativeTask[$rTask]=([double]$vpTask[$rTask+5]-[double]$vmTask[$rTask+5])/(2*[double]$probeTask.h)}
        $fdTask.Add($derivativeTask)
      }
      $probesTask.Add(@{name=$probeTask.name;ok=$okTask;h=$probeTask.h;strain=$probeTask.strain;previous=$probeTask.previous;base_values=$baseTask;fd_columns=$fdTask;branch_probes=$branchTask})
    }
    $recordsTask.Add(@{case=$rowTask[0];phi=$rowTask[1];ey=$rowTask[2];gxy=$rowTask[3];values=$valuesTask;tangent_probes=$probesTask})
  }
  if((Get-FileHash -LiteralPath $SourceWorkbook -Algorithm SHA256).Hash -ne $sourceHashTask){throw 'Source workbook changed'}
  @{source_sha256=$sourceHashTask;selftest_inputs_trial_repair=[bool]$RepairSelfTestInputs;built_in_material_self_test=$selfTask;built_in_edge_self_test=$edgeTask;self_test_table_flat=$selfRowsTask;value_names=@('ok','converged','elastic','failure_code','failure_message','sx','sy','tau','sz','epx','epy','gp','epz','yield_function','plastic_multiplier','t00','t01','t02','t10','t11','t12','t20','t21','t22');records=$recordsTask}|ConvertTo-Json -Depth 8|Set-Content -LiteralPath (Join-Path $OutputRoot 'material_points_raw.json') -Encoding UTF8
}finally{
  if($null -ne $wbTask){$wbTask.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($wbTask)}
  $xlTask.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xlTask)
  [GC]::Collect();[GC]::WaitForPendingFinalizers()
}

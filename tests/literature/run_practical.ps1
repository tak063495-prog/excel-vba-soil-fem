param(
  [string]$Names='*',
  [ValidateSet('INCONSISTENT','DAVIS')][string]$FlowPolicy='INCONSISTENT',
  [switch]$AsControl,
  [string]$PythonPath='python',
  [string]$SourceWorkbook=(Join-Path $PSScriptRoot '../../workbook/2DSoilFEM_20261008_practical.xlsm'),
  [string]$OutputRoot=(Join-Path $PSScriptRoot '../tmp/practical_validation')
)
$ErrorActionPreference='Stop'
$protocolTask=Get-Content (Join-Path $PSScriptRoot 'practical_protocol.json') -Raw -Encoding UTF8|ConvertFrom-Json
$sourceTask=(Resolve-Path -LiteralPath $SourceWorkbook).Path
$shaTask=(Get-FileHash -LiteralPath $sourceTask -Algorithm SHA256).Hash
$reviewedHashesTask=@($protocolTask.source_sha256)+@($protocolTask.additional_reviewed_sources|ForEach-Object {$_.sha256})
if($shaTask.ToLowerInvariant() -notin $reviewedHashesTask){throw 'Workbook source hash differs from the reviewed protocol; update the protocol for a new build before comparing results.'}
$buildTask=if($shaTask.ToLowerInvariant() -eq $protocolTask.source_sha256){$protocolTask.source_build}else{($protocolTask.additional_reviewed_sources|Where-Object {$_.sha256 -eq $shaTask.ToLowerInvariant()}).build}
$suffixTask='_critical'+($buildTask -replace '^.*INCO_','')
if($FlowPolicy -cne $protocolTask.controls.flow_policy -and -not $AsControl){throw 'Different flow policy requires -AsControl; DAVIS changes the frictional material/flow comparison.'}
$outputTask=[IO.Path]::GetFullPath($OutputRoot)
New-Item -ItemType Directory -Force -Path $outputTask|Out-Null
$fixturesTask=Join-Path $outputTask 'fixtures.json'
& $PythonPath (Join-Path $PSScriptRoot 'build_practical_fixtures.py') --output $fixturesTask
if($LASTEXITCODE -ne 0){throw 'Fixture generation failed'}
$dataTask=Get-Content -LiteralPath $fixturesTask -Raw -Encoding UTF8|ConvertFrom-Json
$selectedTask=@($dataTask.fixtures|Where-Object {$_.name -like $Names})
if($selectedTask.Count -eq 0){throw 'No matching practical verification cases'}
$runnerShaTask=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'run_cases.ps1') -Algorithm SHA256).Hash
$fixtureShaTask=(Get-FileHash -LiteralPath $fixturesTask -Algorithm SHA256).Hash
@{source_sha256=$shaTask;fixture_sha256=$fixtureShaTask;runner_sha256=$runnerShaTask;flow_policy=$FlowPolicy;as_control=[bool]$AsControl;cases=@($selectedTask.name);started_at=(Get-Date).ToString('o')}|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $outputTask 'invocation.json') -Encoding UTF8
foreach($caseTask in $selectedTask){
  $folderTask=Join-Path $outputTask ($caseTask.name+$(if($FlowPolicy -eq 'DAVIS'){'_davis_control'}else{''}))
  & (Join-Path $PSScriptRoot 'run_cases.ps1') -Mode SRM -Names $caseTask.name -FlowPolicy $FlowPolicy -FixturePath $fixturesTask -CaseSuffix $suffixTask -SrmTolerance $protocolTask.controls.srm_tolerance -SrmMaximum $protocolTask.controls.srm_fmax -SourceWorkbook $sourceTask -OutputRoot $folderTask
  if((Get-FileHash -LiteralPath $sourceTask -Algorithm SHA256).Hash -cne $shaTask){throw 'Source workbook changed during verification'}
}
& $PythonPath (Join-Path $PSScriptRoot 'analyze_practical_results.py') --output $outputTask --fixtures $fixturesTask
if($LASTEXITCODE -ne 0){throw 'Evidence controls/hash audit failed'}

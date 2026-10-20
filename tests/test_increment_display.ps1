param(
    [Parameter(Mandatory=$true)][string]$Book,
    [string]$Out = 'tests/results/increment_display.json'
)
$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
Push-Location $repo
$app = $null; $wb = $null; $checks = [Collections.Generic.List[string]]::new()
try {
    if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Another Excel process is running; close Excel before this test.' }
    $source = (Resolve-Path -LiteralPath $Book).Path
    $sourceHash = (Get-FileHash -LiteralPath $source).Hash
    $tmp = Join-Path $repo ('tests/tmp/increment_display_' + [Guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp | Out-Null
    $copy = Join-Path $tmp 'increment_display.xlsm'
    Copy-Item -LiteralPath $source -Destination $copy
    $dll = Join-Path (Split-Path $source) 'soilfem_native_x64.dll'
    if (Test-Path -LiteralPath $dll) { Copy-Item -LiteralPath $dll -Destination (Join-Path $tmp 'soilfem_native_x64.dll') }

    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false; $app.DisplayAlerts = $false; $app.EnableEvents = $false; $app.AutomationSecurity = 1
    $wb = $app.Workbooks.Open($copy, 0, $false)
    $cm = $wb.VBProject.VBComponents.Add(1)
    $cm.Name = 'IncrementDisplayHarness'
    $cm.CodeModule.AddFromString((Get-Content -LiteralPath (Join-Path $repo 'tests/native/IncrementDisplayHarness.bas') -Raw -Encoding UTF8))
    $prefix = "'" + $wb.Name + "'!"
    $compile = $app.VBE.CommandBars.FindControl(1, 578)
    if ($compile.Enabled) { $compile.Execute() }
    if ($app.VBE.CommandBars.FindControl(1, 578).Enabled) { throw 'VBA compile failed.' }
    $checks.Add('compile')

    function Check([bool]$ok, [string]$label) { if (-not $ok) { throw "FAIL $label" }; $checks.Add($label) }
    function Run([string]$name, [object[]]$macroArgs = @()) {
        if ($macroArgs.Count -eq 0) { return $app.Run($prefix + $name) }
        if ($macroArgs.Count -eq 1) { return $app.Run($prefix + $name, $macroArgs[0]) }
        throw 'This harness only uses zero or one macro argument.'
    }
    Check (([string](Run 'TIncrementMath')) -eq 'PASS') 'increment math'
    foreach ($kind in 'zero','constant','linear','axes','shear','negative_shear','thin') {
        $svg = [string](Run 'TIncrementSvgFixture' @($kind))
        Check ($svg.Length -gt 0) "svg $kind nonempty"
        Check ($svg.Contains('increment-overlay')) "svg $kind layer"
        [xml]$doc = $svg
        $paths = @($doc.SelectNodes("//*[local-name()='path' and @data-level]"))
        if($kind -eq 'thin'){
            $line=$doc.SelectSingleNode('//line[@data-principal="max"]')
            Check ([Math]::Abs([double]$line.x2-[double]$line.x1)-gt 0.2) 'one thin element does not collapse all cross lengths'
        }
        if ($kind -eq 'axes') {
            $lines = @($doc.SelectNodes("//*[local-name()='line' and @data-principal]"))
            Check ($lines.Count -eq 16) 'axes 16 principal lines'
            Check (($svg -split "stroke='#d62728'").Count -gt 1) 'axes tension color'
            Check (($svg -split "stroke='#1f77b4'").Count -gt 1) 'axes compression color'
        }
        if ($kind -eq 'zero' -or $kind -eq 'constant') { Check ($paths.Count -eq 0) "$kind no contours" }
        if ($kind -eq 'linear') { Check ($paths.Count -gt 0) 'linear contour paths' }
        if ($kind -eq 'linear') {
            foreach($path in $paths){
                $xy=@([regex]::Matches($path.d,'-?\d+(?:\.\d+)?')|ForEach-Object{[double]$_.Value})
                $want=([double]$path.'data-level'-0.01)/0.005
                Check ([Math]::Abs($xy[0]+$xy[1]-$want) -lt 1e-5 -and [Math]::Abs($xy[2]+$xy[3]-$want) -lt 1e-5) 'analytic linear contour endpoints'
            }
        }
        if($kind -in 'axes','shear','negative_shear'){
            foreach($line in @($doc.SelectNodes('//line[@data-principal]'))){
                $dx=[double]$line.x2-[double]$line.x1;$dy=[double]$line.y2-[double]$line.y1
                if($kind -eq 'axes'){
                    if($line.'data-principal' -eq 'max'){Check ($line.stroke -eq '#d62728' -and [Math]::Abs($dy) -lt 1e-8 -and [Math]::Abs([Math]::Abs($dx)-0.3) -lt 1e-8) 'tension full-length horizontal'}
                    else{Check ($line.stroke -eq '#1f77b4' -and [Math]::Abs($dx) -lt 1e-8 -and [Math]::Abs([Math]::Abs($dy)-0.15) -lt 1e-8) 'compression half-length vertical'}
                }elseif($line.'data-principal' -eq 'max'){
                    $want=if($kind -eq 'shear'){-1.0}else{1.0}
                    Check ([Math]::Abs($dy/$dx-$want) -lt 1e-5) "$kind correct Y-flipped direction"
                }
            }
        }
    }

    $ws = $wb.Worksheets.Item('出力・表示')
    Check ($ws.Range('B28').Font.Size -eq 11) 'readable shear checkbox cell label'
    Check ($ws.Range('B29').Font.Size -eq 11) 'readable plastic checkbox cell label'
    Check ($ws.Range('C18').HorizontalAlignment -eq -4131) 'content selector left aligned'
    for ($id = 1; $id -le 70; $id++) {
        $label = [string](Run 'FEMResultChoiceLabel' @($id))
        Check ($label.Length -gt 0) "label $id"
        Check (([int](Run 'FEMResultChoiceId' @($label))) -eq $id) "named selector $id"
        Check (([int](Run 'FEMResultChoiceId' @($id))) -eq $id) "legacy selector $id"
    }
    $ws.Range('C27').Value2 = 'A：選択結果へ等値線を追加'
    $ws.Range('J28').Value2 = $true; $ws.Range('J29').Value2 = $false
    Check ([bool](Run 'FEMShearOverlayOn')) 'shear checkbox on'
    Check (-not [bool](Run 'FEMPlasticCrossOn')) 'plastic checkbox independent off'
    $ws.Range('J28').Value2 = $false; $ws.Range('J29').Value2 = $true
    Check (-not [bool](Run 'FEMShearOverlayOn')) 'shear checkbox independent off'
    Check ([bool](Run 'FEMPlasticCrossOn')) 'plastic checkbox on'
    Run 'FEMResultOptionsInstall' | Out-Null
    Check ([bool]$ws.Range('J29').Value2) 'plastic checkbox survives refresh'
    Check (-not [bool]$ws.Range('J28').Value2) 'shear checkbox survives refresh'
    Check (([string](Run 'FEMResultChoiceId' @('invalid-result'))) -eq '0') 'invalid selector rejected'
    $ws.Range('C27').Value2 = 'invalid-mode'
    Check (([string](Run 'TOptionError' @('mode'))).StartsWith('ERROR|')) 'invalid mode rejected'
    $ws.Range('C27').Value2 = 'B：増分専用カラー'
    Check (-not [bool](Run 'FEMOverlayModeA')) 'B mode'
    $ws.Range('C30').Value2 = 1.5
    Check ([double](Run 'FEMPlasticCrossScale') -eq 1.5) 'scale accepted'
    $ws.Range('C30').Value2 = 0
    Check (([string](Run 'TOptionError' @('scale'))).StartsWith('ERROR|')) 'invalid scale rejected'
    $ws.Range('C30').Value2 = 1.5
    $wb.Save()
    $wb.Close($true); $wb = $null
    $wb = $app.Workbooks.Open($copy, 0, $false)
    $ws = $wb.Worksheets.Item('出力・表示')
    Check ([bool]$ws.Range('J29').Value2) 'plastic checkbox persisted'
    Check (-not [bool]$ws.Range('J28').Value2) 'shear checkbox persisted independently'
    Check ((Get-FileHash -LiteralPath $source).Hash -eq $sourceHash) 'source unchanged'
    $result = [ordered]@{ status='PASS'; checks=$checks; book=$copy; sourceHash=$sourceHash }
    $outPath = Join-Path $repo $Out; New-Item -ItemType Directory -Force -Path (Split-Path $outPath) | Out-Null
    $result | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $outPath -Encoding UTF8
    $result | ConvertTo-Json -Depth 4
}
finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    if ($app) { try { $app.Quit() } catch {} }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    Pop-Location
}

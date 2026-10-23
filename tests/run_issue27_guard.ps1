param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no dispatch.'}
$Book=(Resolve-Path $Book).Path;$hash=(Get-FileHash $Book).Hash;$Out=[IO.Path]::GetFullPath($Out)
if(Test-Path $Out){throw 'Preserve previous evidence'}
New-Item -ItemType Directory $Out|Out-Null;$copy=Join-Path $Out 'guard.xlsm';Copy-Item $Book $copy
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1;$wb=$xl.Workbooks.Open($copy,0,$false)
 $m=$wb.VBProject.VBComponents.Add(1);$m.Name='Issue27Guard'
 $m.CodeModule.AddFromString(@'
Public Function I27Check() As String
 On Error GoTo Failed
 FEMUiPreflightRuntimeLayout ThisWorkbook.Worksheets("高速化設定")
 I27Check="PASS"
 Exit Function
Failed: I27Check="ERROR|" & Err.Number & "|" & Err.Description
End Function
'@)
 $name=$wb.Name;$ws=$wb.Worksheets.Item('高速化設定')
 $ws.Range('C12').Formula='=C11'
 $normal=$xl.Run("'$name'!I27Check");if($normal -ne 'PASS'){throw $normal}
 # Test copy only: simulate a pre-existing legacy layout; current UI protects rows.
 $ws.Unprotect()
 $null=$ws.Range('A9').EntireRow.Insert()
 $before=ConvertTo-Json -InputObject $ws.UsedRange.Formula -Depth 5 -Compress
 $guard=$xl.Run("'$name'!I27Check")
 if($guard -notlike 'ERROR|*' -or $guard -notmatch '数式'){throw "Formula relocation was not blocked: $guard"}
 $after=ConvertTo-Json -InputObject $ws.UsedRange.Formula -Depth 5 -Compress
 if($before -ne $after){throw 'Guard changed settings/formulas/layout'}
 $wb.Save();$wb.Close($false);$wb=$xl.Workbooks.Open($copy,0,$true)
 $ws=$wb.Worksheets.Item('高速化設定')
 if((ConvertTo-Json -InputObject $ws.UsedRange.Formula -Depth 5 -Compress) -ne $before){throw 'Saved/reopened guard state differs'}
 @{status='PASS';normal=$normal;guard=$guard;settings_unchanged=$true;reopen=$true;book_sha256=$hash}|ConvertTo-Json|Set-Content (Join-Path $Out 'guard.json') -Encoding utf8
 $wb.Close($false)
}finally{if($wb){try{$wb.Close($false)}catch{}};$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl);[GC]::Collect();[GC]::WaitForPendingFinalizers()}
if((Get-FileHash $Book).Hash -ne $hash){throw 'Source workbook changed'}

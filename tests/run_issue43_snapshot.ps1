param([Parameter(Mandatory=$true)][string]$Book,[Parameter(Mandatory=$true)][string]$Out,[string]$DirtyProfile='W',[int]$DeactivateMaterial=0,[switch]$RequireTransientFlow)
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel is running; no dispatch.'}
$Book=(Resolve-Path $Book).Path;$Out=[IO.Path]::GetFullPath($Out)
if(Test-Path $Out){throw 'Preserve previous evidence.'}
New-Item -ItemType Directory $Out|Out-Null
$hash=(Get-FileHash $Book).Hash;$copy=Join-Path $Out 'snapshot_test.xlsm'
Copy-Item $Book $copy;Copy-Item (Join-Path (Split-Path $Book) 'soilfem_native_x64.dll') $Out
$xl=New-Object -ComObject Excel.Application
try{
 $xl.DisplayAlerts=$false;$xl.EnableEvents=$false;$xl.AutomationSecurity=1;$wb=$xl.Workbooks.Open($copy,0,$false)
 $engine=$wb.VBProject.VBComponents.Item('FEMEngine').CodeModule
 # Wrappers live in the same test-copy module so the real private SRM routines
 # are exercised without adding public debug entry points to the distribution.
 $engine.AddFromString(@'
Public Function I43Run() As String
 On Error GoTo Failed
 SuppressUserMessages=True
 FEMWriteTextSetting "EXPORT_LOAD", ""
 FEMWriteTextSetting "SRM_FIXED_FS", "1"
 P0_RunAnalysis
 I43Run=ResultStatus & "|" & AnalysisMessage
 Exit Function
Failed: I43Run="ERROR|" & Err.Description
End Function
Public Sub I43Save()
 P3CaptureSrmSuccessSnapshot P3CurrentStrengthFactor
 If Not P3SrmSnapReady Then Err.Raise vbObjectError+8200,,"No accepted snapshot"
End Sub
Public Function I43Dirty(ByVal profile As String, ByVal materialId As Long) As String
 On Error GoTo Failed
 If materialId>0 Then
  P3SetMaterialElementsActive materialId,False
  P3ActiveSetGen=P3ActiveSetGen+1
  P6InvalidateActiveDependentCaches
 End If
 GWStage profile
 GWUseFinalPressure
 I43Dirty="DIRTY"
 Exit Function
Failed: I43Dirty="DIRTY_ERROR|" & Err.Description
End Function
Public Function I43Restore() As Boolean
 I43Restore=P3RestoreSrmSuccessSnapshot()
End Function
'@)
 $water=$wb.VBProject.VBComponents.Item('FEMWater').CodeModule
 $water.AddFromString(@'
Public Sub I43Dump(ByVal path As String)
 Dim f As Integer,e As Long,g As Long,n As Long,item As Variant
 Dim qx As Double,qy As Double,h As Double,prescribed As Boolean,available As Boolean
 f=FreeFile: Open path For Output As #f
 Print #f,"META," & mMode & "," & mProfile & "," & mOverrideProfile & "," & mSource & "," & mSig
 Print #f,"FLAGS," & GWActive & "," & mReady & "," & mFlowSolved & "," & GJReady
 Print #f,"DIAG," & mRel & "," & mBalance & "," & mSeconds & "," & mLocalError & "," & mFluxCorrection & "," & mInflow & "," & mOutflow
 For e=0 To NumberOfElement-1
  Print #f,"ACTIVE," & e & "," & P3IsElementActive(e) & "," & mK(e)
  For g=0 To 3
   qx=0#: qy=0#: available=GWViewFlux(e,g,qx,qy)
   Print #f,"VIEWFLUX," & e & "," & g & "," & available & "," & qx & "," & qy
   Print #f,"P," & e & "," & g & "," & mTarget(g,e) & "," & mCurrent(g,e) & "," & mCommitted(g,e) & "," & mStart(g,e)
  Next g
  For g=0 To 7
   h=0#: prescribed=False: available=GWViewHead(e,g,h,prescribed)
   Print #f,"VIEWHEAD," & e & "," & g & "," & available & "," & h & "," & prescribed
   Print #f,"ALIAS," & e & "," & g & "," & GWFlowNode(e,g) & "," & GJTopoNode(Elem(e).node(g))
  Next g
 Next e
 For n=0 To lastDof
  Print #f,"DOF," & n & "," & TDisp(n) & "," & P3FinalInternalForce(n) & "," & mWaterForce(n) & "," & P3AppliedForce(n) & "," & P3SelfWeightForce(n) & "," & NodeCond(n)
 Next n
 If mFlowSolved Then
  For n=0 To NumberOfFreeNode-1
   Print #f,"HEAD," & n & "," & mHeads(n) & "," & mHeadFixed(n) & "," & mFlowReaction(n)
  Next n
  For e=0 To NumberOfElement-1
   For g=0 To 3
    Print #f,"FLUX," & e & "," & g & "," & mFx(g,e) & "," & mFy(g,e)
    If Not FSActive Then Print #f,"FACE," & e & "," & g & "," & mFaceFlows(g,e)
   Next g
  Next e
 End If
 For Each item In mSurfaceRows
  Print #f,"SURFACE," & item(0) & "," & item(1) & "," & item(2) & "," & item(3) & "," & item(4) & "," & item(5) & "," & item(6)
 Next item
 Close #f
End Sub
'@)
 $ctl=$xl.VBE.CommandBars.FindControl(1,578);if($ctl.Enabled){$ctl.Execute()}
 if($xl.VBE.CommandBars.FindControl(1,578).Enabled){throw 'Snapshot test driver failed compile'}
 $name=$wb.Name;$status=$xl.Run("'$name'!I43Run");if($status -notlike 'PASS|*'){throw $status}
 $null=$xl.Run("'$name'!I43Save")
 $null=$xl.Run("'$name'!I43Dump",(Join-Path $Out 'accepted.csv'))
 $dirty=$xl.Run("'$name'!I43Dirty",$DirtyProfile,$DeactivateMaterial)
 if($dirty -ne 'DIRTY'){throw "Dirty-state setup failed: $dirty"}
 $null=$xl.Run("'$name'!I43Dump",(Join-Path $Out 'last_trial.csv'))
 if((Get-Content (Join-Path $Out 'accepted.csv') -Raw) -eq (Get-Content (Join-Path $Out 'last_trial.csv') -Raw)){throw 'Trial did not change accepted water/topology'}
 if(-not $xl.Run("'$name'!I43Restore")){throw 'Accepted snapshot restoration failed'}
 $null=$xl.Run("'$name'!I43Dump",(Join-Path $Out 'restored.csv'))
 if((Get-Content (Join-Path $Out 'accepted.csv') -Raw) -ne (Get-Content (Join-Path $Out 'restored.csv') -Raw)){throw 'Restored water/mechanics differ from accepted snapshot'}
 if($RequireTransientFlow){if(-not (Select-String -Path (Join-Path $Out 'restored.csv') -Pattern '^VIEWFLUX,[0-9]+,[0-9]+,True,') -or -not (Select-String -Path (Join-Path $Out 'restored.csv') -Pattern '^VIEWHEAD,[0-9]+,[0-9]+,True,')){throw 'Transient flow/head unavailable after SRM restore'}}
 @{status='PASS';require_transient_flow=$RequireTransientFlow.IsPresent;dirty=$dirty;deactivated_material=$DeactivateMaterial;book_sha256=$hash;scope='actual private SRM save/restore with deliberately changed later-trial water/active topology; no claim that the dirty state was generated by natural mechanical nonconvergence'}|ConvertTo-Json|Set-Content (Join-Path $Out 'snapshot.json') -Encoding utf8
 $wb.Close($false)
}finally{if($wb){try{$wb.Close($false)}catch{}};$xl.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl);[GC]::Collect();[GC]::WaitForPendingFinalizers()}
if((Get-FileHash $Book).Hash -ne $hash){throw 'Source evidence workbook was changed'}

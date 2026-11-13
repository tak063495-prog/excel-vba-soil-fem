Public Function TrialTangentRestoreAudit() As Variant
 Dim k As Long,m As Long,e As Long,g As Long,i As Long,j As Long,n As Long,p As Long
 Dim base() As Double,internal() As Double,plasticK(0 To 15,0 To 15) As Double
 Dim expected As Element_Data,delta As Double,matrixError As Double,plasticCount As Long,elasticCount As Long,dirty As Boolean,status As String
 SuppressUserMessages=True:P0_RunAnalysis:status=ResultStatus
 If status<>"PASS" Then TrialTangentRestoreAudit=Array(status,0,0,False,0#,0#,0):Exit Function
 e=-1
 For k=0 To NumberOfElement-1
  If P6MaterialKindOf(Elem(k).MatNo)="SOIL" Then e=k:Exit For
 Next k
 If e<0 Then Err.Raise 1009,,"SOIL fixture required"
 m=Elem(e).MatNo
 Material(m).fai=30#:Material(m).psai=30#:Material(m).cohesion=1#
 Material(m).SinFriction=Sin(30#*3.14159265358979/180#):Material(m).CosFriction=Cos(30#*3.14159265358979/180#)
 Material(m).SinDilation=Material(m).SinFriction:Material(m).CosDilation=Material(m).CosFriction
 For k=0 To NumberOfElement-1
  For g=0 To 3
   For i=0 To 16:Elem(k).mStmat(i,g)=0#:Next i
   P3CommittedYielded(g,k)=False:Elem(k).P2Yielded(g)=False
   P3CommittedSpmatValid(g,k)=True
   For i=0 To 2
    For j=0 To 15:P3CommittedSpmat(i,j,g,k)=Elem(k).Smat(i,j,g):Next j
    For j=0 To 2:P3CommittedDpmat(i,j,g,k)=Elem(k).Dmat(i,j):Next j
   Next i
  Next g
 Next k
 ReDim base(lastDof):ReDim internal(lastDof)
 For i=0 To lastDof Step 2:TDisp(i)=-0.02*XXX(i)+0.01*XXX(i+1):TDisp(i+1)=0.019*XXX(i+1):Next i
 If Not P3EvaluateTrialState(base,internal) Then Err.Raise 1010,,"Plastic trial failed"
 For g=0 To 3:If Elem(e).P2Yielded(g) Then plasticCount=plasticCount+1
 Next g
 P3MarkAllTangentDirty
 If Not P3RebuildTangentFromSpmat() Then Err.Raise 1011,,"Plastic tangent failed"
 For i=0 To 15:For j=0 To 15:plasticK(i,j)=Elem(e).kmat(i,j):Next j:Next i
 P3RestoreCommittedMaterialState False
 dirty=Elem(e).TangentDirty
 For i=0 To lastDof Step 2:TDisp(i)=-0.0000001*XXX(i):TDisp(i+1)=-0.0000001*XXX(i+1):Next i
 If Not P3EvaluateTrialState(base,internal) Then Err.Raise 1012,,"Elastic trial failed"
 For g=0 To 3:If Elem(e).P2Yielded(g) Then elasticCount=elasticCount+1
 Next g
 If Not P3RebuildTangentFromSpmat() Then Err.Raise 1013,,"Elastic tangent failed"
 expected=Elem(e)
 SetElmStiffness Material(m).thickness,Material(m).weight,expected.bMat,expected.Smat,expected.kmat,expected.ElmWeight,expected.dj,expected.n
 If Not(P3SrmEnabled And P3FlowPolicyIsInconsistent() And Not P3IncoElasticCorrection) Then P6AddQ8HourglassStabilization expected,Material(m).thickness,True
 For i=0 To 15
  For j=0 To 15
   If Abs(plasticK(i,j)-expected.kmat(i,j))>delta Then delta=Abs(plasticK(i,j)-expected.kmat(i,j))
   If Abs(Elem(e).kmat(i,j)-expected.kmat(i,j))>matrixError Then matrixError=Abs(Elem(e).kmat(i,j)-expected.kmat(i,j))
  Next j
 Next i
 TrialTangentRestoreAudit=Array(status,plasticCount,elasticCount,dirty,delta,matrixError,FEMNativeStateCopies)
End Function

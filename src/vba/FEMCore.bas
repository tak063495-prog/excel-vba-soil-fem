Option Explicit
' Shared types, constants, and analysis state. Other FEM* modules read/write these Public names.

'
' 2DSoilFEM (plane-strain Q8, 2x2 Gauss, Mohr-Coulomb)
' ----------------------------------------------------
' Input sheets: 設定 / 材料データ / ステージ / 載荷 / 接合 / 節点データ / 要素データ
' Help sheet : 仕様  (Japanese option list; regenerated on open)
' Run order  : GRAVITY / LOAD / UNLOAD / BIRTH / DEATH / MATSET / RESET_* / SRM
' Checkpoints: *_out/stgNN_KIND/{meta.txt,node.csv,gauss.csv,dof.csv,origin.csv}
' Resume     : 設定 EXPORT_LOAD = empty | AUTO | folder name
'
' Prefixes: P1 results, P2 return mapping, P3 staging/SRM, P5 mesh quality, P6 solver
' RCM reorders internal DOFs; CSV I/O always uses original node/DOF ids.
'
Public Type Material_Data '材料データ
  Young As Double         'ヤング率
  Poisson As Double       'ポアソン比
  thickness As Double     '厚さ
  weight As Double       '重さ
  fai As Double
  psai As Double
  cohesion As Double
  ReduceStrength As Boolean 'SRMでc・φを下げるか
  InitialActive As Boolean  'ON=解析開始から有効、OFF=BIRTH待ち
  kind As String          'SOIL / STRUCT / JOINT
  kn As Double            'JOINT 法線剛性
  ks As Double            'JOINT せん断剛性
  AllowTension As Boolean 'JOINT False=剥離可
  ElasticD00 As Double    '平面ひずみ弾性係数の材料別キャッシュ
  ElasticD01 As Double
  ElasticD22 As Double
  SinFriction As Double   '摩擦角・ダイレタンシー角の三角関数キャッシュ
  CosFriction As Double
  SinDilation As Double
  CosDilation As Double
  MaterialCacheReady As Boolean
End Type

'P2材料点の入出力。Excelシート・全体変位配列を参照しない。
Public Type P2_MaterialPointInput
  PreviousStress(3) As Double '確定応力 sx, sy, txy, sz。添え字3はσz
  StrainIncrement(2) As Double '今回の工学ひずみ増分 ex, ey, gxy。εz=0
  Young As Double
  Poisson As Double
  frictionAngle As Double '度
  cohesion As Double
  dilationAngle As Double '度
  tolerance As Double
  maxIterations As Long
  maxSubsteps As Long
  EnableSubstepping As Boolean
  PreferredSubsteps As Long
  UseCachedConstants As Boolean
  CachedElasticD00 As Double
  CachedElasticD01 As Double
  CachedElasticD22 As Double
  CachedSinFriction As Double
  CachedCosFriction As Double
  CachedSinDilation As Double
  CachedCosDilation As Double
End Type

Public Type P2_MaterialPointOutput
  TrialStress(3) As Double 'sx, sy, txy, sz
  Stress(3) As Double 'sx, sy, txy, sz
  Tangent(2, 2) As Double
  PlasticStrain(3) As Double 'ex, ey, gxy, ez
  PlasticMultiplier As Double
  TrialYieldFunction As Double
  YieldFunction As Double
  PrincipalStress(2) As Double '最大・中間・最小主応力
  principalAngle As Double '度
  yielded As Boolean
  PlasticOccurred As Boolean
  converged As Boolean
  Elastic As Boolean
  failureCode As Long
  failureMessage As String
  iterations As Long
  UsedSubsteps As Long
End Type

Public Type Element_Data '要素データ
  node(7) As Long '節点番号　添え字0～7
  x(7) As Double '節点X座標　添え字0～7
  y(7) As Double '節点Y座標　添え字0～7
  u(15) As Double '自由度×8個
  MatNo As Long  '材料番号
  area As Double    '面積
  ElNode(15) As Long '自由度×8個
  Bmat(2, 15, 3) As Double 'B行列
  Smat(2, 15, 3) As Double 'S行列
  Spmat(2, 15, 3) As Double 'S行列
  Mfai As Double
  Mcohesion As Double
  Mpsai As Double
  p(2) As Double
  ElNo As Long
  Stmat(16, 3) As Double '応力状態。16は平面ひずみの面外応力σz
  dStmat(2, 3) As Double
  
  mStmat(16, 3) As Double '前回応力状態。16はσz
  P2PlasticStrain(3, 3) As Double 'P2直前の局所塑性ひずみ増分 ex, ey, gxy, ez
  P2PlasticMultiplier(3) As Double 'P2直前の局所塑性乗数
  P2YieldFunction(3) As Double 'P2降伏関数値
  P2Yielded(3) As Boolean 'P2塑性状態フラグ
  P2PrincipalStress(2, 3) As Double 'P2更新後の最大・中間・最小主応力
  P2LastSuccessfulSubsteps(3) As Long
  P2EasySubstepStreak(3) As Long
  
  Dmat(2, 2) As Double '弾性係数行列
  Dpmat(2, 2, 3) As Double '4Gauss点の平面ひずみ接線係数行列
  kmat(15, 15) As Double '要素剛性行列
  RN(3, 15) As Double
  n(3, 7) As Double
  dj(3) As Double
  ElmWeight As Double '自重
  HourglassModeCount As Long
  HourglassScale As Double
  HourglassMode(3, 15) As Double '2×2核のゼロエネルギーモード（最大4）
  HourglassShapeGen As Long
  IsJoint As Boolean
  SpmatValid(3) As Boolean 'Gauss点接線が今回の構成則で書いたものか。頂点ゼロ接線もTrue
  JointContact(3) As Long 'JOINT Gauss: -1未設定 0密着 1剥離 2滑り
  TangentDirty As Boolean
  HourglassKReady As Boolean
  HourglassK(15, 15) As Double
  HourglassKScaledReady As Boolean
  HourglassCachedScale As Double
  HourglassKScaled(15, 15) As Double
End Type
Public nx As Long
Public my As Long
Public NumberOfFreeNode As Long '計算節点データ数
Public NumberOfNode As Long   '節点データ数
Public NumberOfMaterial As Long '材料データ数
Public NumberOfElement As Long '要素データ数
Public FEMLastProc As String
Public nn As Long '要素配列の最終添え字
Public NAB As Long '節点配列の最終添え字
Public x() As Double '節点X座標
Public y() As Double  '節点Y座標
Public XXX() As Double
Public Force() As Double '節点外力
Public Disp() As Double  '節点変位
Public nmn As Long
Public ESRR As Double
Public TDisp() As Double  '節点変位




Public RForce() As Double '節点外力
Public NForce() As Double '節点外力
Public eForce() As Double '節点外力
Public UDisp() As Double  '節点変位

Public iNForce() As Double '節点外力
Public mRisp() As Double 'RCMで並べ替える既存補助配列

' Q8入力の節点番号を、解析用の自由節点番号へ対応付ける配列。
Public FEMFreeNodeMap() As Long
Public nh As Double
Public nv As Double
Public FSS As Double
Public mode As Long
Public NodeCond() As Long '変位境界条件

Public Material() As Material_Data '材料データ
Public Elem() As Element_Data '要素データ

Public kmat() As Double '要素剛性行列
Public TotalMat() As Double '全体剛性行列（混合u-pのSchur後）
Public TotalMat2() As Double '互換用。BAND経路では使わない
Public OriginalMat() As Double '境界条件適用前の剛性行列
Public OriginalForce() As Double '境界条件適用前の外力
Public Reaction() As Double '残差・反力
Public BandWidth As Long 'バンド幅

Public Const RESULT_NOT_RUN As String = "NOT_RUN"
Public Const RESULT_PASS As String = "PASS"
Public Const RESULT_INPUT_ERROR As String = "INPUT_ERROR"
Public Const RESULT_MATERIAL_ERROR As String = "MATERIAL_ERROR"
Public Const RESULT_NONCONVERGED As String = "NONCONVERGED"
Public Const RESULT_GLOBAL_SINGULAR As String = "GLOBAL_SINGULAR"
Public Const RESULT_CAPACITY_ERROR As String = "CAPACITY_ERROR"
Public Const RESULT_RUNTIME_ERROR As String = "RUNTIME_ERROR"

Public Enum NewtonPolicy
  NP_NORMAL = 0
  NP_ROBUST_SRM = 1
End Enum

Public Enum NumericalStatus
  NS_UNKNOWN = 0
  NS_CONVERGED = 1
  NS_FAILED = 2
End Enum

Public Enum MechanicalStatus
  MS_UNKNOWN = 0
  MS_STABLE = 1
  MS_LIMIT_STATE = 2
End Enum

Public Type SRMTrialResult
  Fs As Double
  NumericalStatus As NumericalStatus
  MechanicalStatus As MechanicalStatus
  IncrementCount As Long
  NewtonCount As Long
  MaxDisp As Double
  MaxDispOverH As Double
  ComplianceMaxRatio As Double
  FactorCount As Long
  FactorReuseCount As Long
  ElapsedSec As Double
End Type

Public Type LimitStateMonitor
  RefValue(1 To 3) As Double
  RefCount As Long
  ComplianceRef As Double
  RefReady As Boolean
  RatioCurrent As Double
  RatioMax As Double
  WarningCount As Long
  ConsecutiveTrip As Long
  candidate As Boolean
  FirstCandidateInc As Long
  FirstCandidateLambda As Double
  FirstCandidateUmax As Double
End Type

Public P3NewtonPolicy As NewtonPolicy
Public P3Ls50Enabled As Boolean
Public P6BandBenchOnly As Boolean
Public P6LimitMonitorOn As Boolean
Public P6Limit As LimitStateMonitor
Public P3LastTrial As SRMTrialResult
Public P3FosPass As Double
Public P3FosFail As Double
Public P3FosMid As Double
Public P3FosWidth As Double
Public P3FosBracket As Boolean
Public Const FEM_BUILD_STAMP As String = "20261008_BEST_03"
Public Const P1_FORMULATION_PLANE_STRESS As String = "PLANE_STRESS_2D"
Public Const P1_FORMULATION_PLANE_STRAIN As String = "PLANE_STRAIN_2D"
Public Const P1_RESULT_COUNT As Long = 22
Public Const P1_GAUSS_COUNT As Long = 4
Public Const P1_RESULT_N1_LOCAL As Long = 0
Public Const P1_RESULT_N2_LOCAL As Long = 1
Public Const P1_RESULT_N12_LOCAL As Long = 2
Public Const P1_RESULT_M1_LOCAL As Long = 3
Public Const P1_RESULT_M2_LOCAL As Long = 4
Public Const P1_RESULT_M12_LOCAL As Long = 5
Public Const P1_RESULT_Q1_LOCAL As Long = 6
Public Const P1_RESULT_Q2_LOCAL As Long = 7
Public Const P1_RESULT_SIGMA1_TOP As Long = 8
Public Const P1_RESULT_SIGMA2_TOP As Long = 9
Public Const P1_RESULT_TAU12_TOP As Long = 10
Public Const P1_RESULT_SIGMA1_BOTTOM As Long = 11
Public Const P1_RESULT_SIGMA2_BOTTOM As Long = 12
Public Const P1_RESULT_TAU12_BOTTOM As Long = 13
Public Const P1_RESULT_PRINCIPAL_ANGLE_TOP As Long = 14
Public Const P1_RESULT_VON_MISES_TOP As Long = 15
Public Const P1_RESULT_PRINCIPAL_ANGLE_BOTTOM As Long = 16
Public Const P1_RESULT_VON_MISES_BOTTOM As Long = 17
Public Const P1_RESULT_SIGMA_Z As Long = 18
Public Const P1_RESULT_SIGMA_MAX_3D As Long = 19
Public Const P1_RESULT_SIGMA_MID_3D As Long = 20
Public Const P1_RESULT_SIGMA_MIN_3D As Long = 21
Public Const P2_DEFAULT_TOLERANCE As Double = 0.0000000001
Public Const P2_DEFAULT_MAX_ITERATIONS As Long = 25
Public Const P2_DEFAULT_MAX_SUBSTEPS As Long = 8
Public Const FEM_INTEGRATED_RESULT_SHEET As String = "診断"
Public Const FEM_DIAGNOSTIC_LAST_ROW As Long = 141
Public Const FEM_INTEGRATED_P1_START_ROW As Long = 145
Public Const FEM_INTEGRATED_P2_TEST_START_ROW As Long = 1
Public Const FEM_INTEGRATED_P2_TEST_START_COLUMN As Long = 40
Public Const FEM_INTEGRATED_P2_TEST_END_ROW As Long = 50
Public Const FEM_INTEGRATED_P2_RESULT_START_ROW As Long = 100
Public Const FEM_INTEGRATED_P2_RESULT_START_COLUMN As Long = 40
Public Const P2_FAILURE_INVALID_INPUT As Long = 4101
Public Const P2_FAILURE_DENOMINATOR As Long = 4102
Public Const P2_FAILURE_NOT_CONVERGED As Long = 4103
Public Const P2_FAILURE_NONFINITE As Long = 4104
Public Const P2_FAILURE_VERTEX As Long = 4105
Public P1Formulation As String
Public P1ResultReady As Boolean
Public P1ResultModelRevision As Long
Public GaussResult() As Double
Public GaussGlobalResult() As Double
Public GaussWeight() As Double
Public ElementAverageResult() As Double
Public NodalResult() As Double
Public NodalGlobalResult() As Double
Public NodalResultWeight() As Double
Public NodalX() As Double
Public NodalY() As Double
Public ElementNodeResult() As Double
Public DisplacementMagnitude() As Double
Public P1ResultAvailable(0 To P1_RESULT_COUNT - 1) As Boolean
Public P1LastOutputLastRow As Long
Public P2LastOutputLastRow As Long
Public P1OutputGaussRows() As Variant
Public P1OutputNodalRows() As Variant
Public P1OutputElementNodeRows() As Variant
Public P2OutputRows() As Variant
Public P1OutputHeaderGauss() As Variant
Public P1OutputHeaderNodal() As Variant
Public P1OutputHeaderElementNode() As Variant
Public P1OutputHeaderGaussRowCount As Long
Public P1OutputHeaderNodalRowCount As Long
Public P1OutputHeaderElementNodeRowCount As Long
Public P1OutputHeaderGaussColumnCount As Long
Public P1OutputHeaderNodalColumnCount As Long
Public P1OutputHeaderElementNodeColumnCount As Long
Public P1OutputGaussRowCount As Long
Public P1OutputGaussColumnCount As Long
Public P1OutputNodalRowCount As Long
Public P1OutputNodalColumnCount As Long
Public P1OutputElementNodeRowCount As Long
Public P1OutputElementNodeColumnCount As Long
Public P2OutputRowCount As Long
Public P2OutputColumnCount As Long
Public AnalysisOK As Boolean
Public ResultStatus As String
Public AnalysisMessage As String
Public AnalysisErrorNumber As Long
Public FailureElement As Long
Public FailureGaussPoint As Long
Public FailureIncrement As Long
Public FailureIteration As Long
Public ModelRevision As Long
Public ResultRevision As Long
Public AnalysisRunning As Boolean
Public SuppressUserMessages As Boolean
Public MatrixFactored As Boolean
Public CurrentIncrement As Long
Public CurrentIteration As Long
Public nDof As Long
Public lastDof As Long
Public ResidualNormFull As Double
Public ResidualNormFree As Double
Public RelativeResidualFree As Double
Public ForceNormFree As Double
Public MaxAbsResidualFree As Double
Public EnergyError As Double
Public MatrixSymmetryError As Double
Public TotalAppliedLoadX As Double
Public TotalAppliedLoadY As Double
Public TotalSelfWeight As Double
Public P6WeightMode As String
Public P6WeightNote As String
Public MinDetJCorner As Double
Public MaxDetJCorner As Double
Public MinDetJGauss As Double
Public MaxDetJGauss As Double
Public TimeInput As Double
Public TimeElementStiffness As Double
Public TimeAssembly As Double
Public TimeSolve As Double
Public TimePlastic As Double
Public TimeOutput As Double

' P3全体反復状態。P3CommittedDispとmStmatが増分開始時の確定値、TDispとStmatが試行値。
Public P3CommittedDisp() As Double
Public P3BoundaryDisp() As Double
Public P3AppliedForce() As Double
Public P3SelfWeightForce() As Double
Public P3FinalInternalForce() As Double
Public P3CommittedPlasticStrain() As Double
Public P3CommittedPlasticMultiplier() As Double
Public P3CommittedYieldFunction() As Double
Public P3CommittedYielded() As Boolean
Public P3CommittedPrincipalStress() As Double
Public P3CommittedSpmat() As Double
Public P3CommittedDpmat() As Double
Public P3CommittedSpmatValid() As Boolean
Public P3OriginalMaterial() As Material_Data
Public P3UseNonlinearResidual As Boolean
Public P3CurrentStrengthFactor As Double
Public P3SrmEnabled As Boolean
Public P3SrmTrialCount As Long
Public P3SrmLower As Double
Public P3SrmUpper As Double
Public P3SrmNote As String
Public P3StagePlanText As String
Public P3PlanHasGravity As Boolean
Public P3PlanHasApply As Boolean
Public P3GravityCommitted As Boolean
Public P3ActivePrefix As Long
Public P3ElementActive() As Boolean
Public P3ElementActiveReady As Boolean
Public P3LockedSelfWeight() As Double
Public P3LockedApplied() As Double
Public P3SupportNodeCond() As Long
Public P3SupportBoundaryDisp() As Double
Public P3SupportCondReady As Boolean
Public P3ReplayQuiet As Boolean
Public P3SrmTrialRunning As Boolean
Public P3SrmSnapReady As Boolean
Public P3SrmSnapFs As Double
Public P3SrmSnapStmat() As Double
Public P3SrmSnapMStmat() As Double
Public P3SrmSnapJointContact() As Long
Public P3SrmSnapTDisp() As Double
Public P3SrmSnapUDisp() As Double
Public P3SrmSnapCommittedDisp() As Double
Public P3SrmSnapInternal() As Double
Public P3SrmSnapCommittedPlasticStrain() As Double
Public P3SrmSnapCommittedPlasticMultiplier() As Double
Public P3SrmSnapCommittedYieldFunction() As Double
Public P3SrmSnapCommittedYielded() As Boolean
Public P3SrmSnapCommittedPrincipalStress() As Double
Public P3SrmSnapCommittedSpmat() As Double
Public P3SrmSnapCommittedDpmat() As Double
Public P3SrmSnapCommittedSpmatValid() As Boolean
Public P3SrmSnapActivePlastic() As Boolean
Public P3SrmSnapActiveCount As Long
Public P3SrmSnapSuccessInc As Long
Public P3SrmSnapLastInc As Long
Public P3SrmSnapLastLoad As Double
Public P3SrmSnapMaxDisp As Double
Public P3SrmSnapMaxCorr As Double
Public P3SrmSnapRelRes As Double
Public P3SrmSnapResFull As Double
Public P3SrmSnapResFree As Double
Public P3SrmSnapMaxAbsRes As Double
Public P3SrmSnapEnergy As Double
Public P3SrmSnapRx As Double
Public P3SrmSnapRy As Double
Public P3SrmSnapInc As Long
Public P3SrmSnapIter As Long
Public P3SrmSnapPrefix As Long
Public P3SrmSnapActive() As Boolean
Public P3SrmSnapNodeCond() As Long
Public P3SrmSnapBoundaryDisp() As Double
Public P3SrmSnapApplied() As Double
Public P3SrmSnapSelfWeight() As Double
Public P3SrmSnapLockedSelf() As Double
Public P3SrmSnapLockedApplied() As Double
Public P3SrmLogStageNo As Long
Public P3SrmLogKind As String
Public P3SrmReplayLimit As Long
Public P3SearchFatal As Boolean
Public P3UserCancel As Boolean
Public P3CommittedNodeCond() As Long
Public P3CommittedBoundaryDisp() As Double
Public P3CommittedBoundaryReady As Boolean
Public P3StageStartDisp() As Double
Public P3StageStartDispReady As Boolean
Public P3SrmTrialFailNote As String
Public P3RunLogRow As Long
Public P3RunLogStageNo As Long
Public P3RunLogKind As String
Public P3StageN As Long
Public P3SrmReplayStart As Long
' Snapshot of active set / BC / loads at last RESET_U, RESET_STRESS or MATSET.
' Required so a later SRM can replay from that origin after a checkpoint resume.
Public P3OriginSnapReady As Boolean
Public P3OriginElementActive() As Boolean
Public P3OriginNodeCond() As Long
Public P3OriginBoundaryDisp() As Double
Public P3OriginAppliedForce() As Double
Public P3OriginLockedSW() As Double
Public P3OriginLockedAP() As Double
Public P3OriginHasGravity As Boolean
Public P3OriginHasApply As Boolean
Public P3SheetMaterial() As Material_Data
Public P3StageKind() As String
Public P3StageId() As Long
Public P3StageGroup() As String
Public P3StageParam() As String
Public P3StageOn() As Boolean
Public P3OrphanHeld() As Boolean
Public P3StageResultRow As Long
Public P3LastCompletedStage As Long
Public P3LastConvergedLoadFactor As Double
Public P3LastConvergedIncrement As Long
Public P3SuccessfulIncrementCount As Long
Public P3GlobalIterationCount As Long
Public P3PlasticPointUpdateCount As Long
Public P3RetryCount As Long
Public P3DefaultIncrementCount As Long
Public P3MaxTrialDisp As Double
Public P3MaxCorrection As Double
Public P3LastRelativeResidual As Double
Public P3StateNote As String
Public P3MaxBoundaryDispError As Double
Public P3RelativeBoundaryDispError As Double
Public P6ReactionSumX As Double
Public P6ReactionSumY As Double
Public Const P3_DEFAULT_INCREMENT_COUNT As Long = 100
Public Const P3_INITIAL_SELF_WEIGHT_INCREMENT_COUNT As Long = 10
Public Const P3_GRAVITY_STEP_SIZE As Double = 0.1
Public Const P3_APPLY_STEP_SIZE As Double = 0.05
Public Const P3_MAX_GLOBAL_ITERATIONS As Long = 50
Public Const P3_MAX_STEP_RETRIES As Long = 8
Public Const P3_MAX_STRENGTH_RETRIES As Long = 25
Public Const P3_MIN_STEP_FACTOR As Double = 0.000001
Public Const P3_STEP_GROWTH As Double = 2#
Public Const P3_MAX_STEP_SIZE As Double = 0.2
Public Const P3_LINESEARCH_MAX As Long = 4
Public Const P3_RESIDUAL_TOLERANCE As Double = 0.00000001
Public Const P3_INCREMENT_TOLERANCE As Double = 0.00000001
Public Const P3_ENGINEERING_RESIDUAL As Double = 0.00001
Public Const P3_ENGINEERING_CORRECTION As Double = 0.0001
Public Const P3_BOUNDARY_DISP_TOLERANCE As Double = 0.00000001
Public Const P6_ITERATIVE_BAND_FALLBACK As Double = 0.001
Public Const P6_HOURGLASS_FACTOR_DEFAULT As Double = 0.05
Public Const FEM_FIXED_DATA_CAPACITY As Long = 60000

' P5メッシュ品質・頑健性診断。解析結果と同じブック内で保持する。
Public P5ModelScale As Double
Public P5GeometryTolerance As Double
Public P5AreaTolerance As Double
Public P5MinElementArea As Double
Public P5MinJacobian As Double
Public P5MaxAspectRatio As Double
Public P5MinAngleDeg As Double
Public P5MaxAngleDeg As Double
Public P5BoundaryEdgeCount As Long
Public P5SharedEdgeCount As Long
Public P5MaterialInterfaceCount As Long
Public P5IsolatedNodeCount As Long
Public P5ConnectedComponentCount As Long
Public P5BoundarySelfIntersectionCount As Long
Public P5LocateCallCount As Long
Public P5LocateTotalIterations As Long
Public P5LocateMaxIterations As Long
Public P5LocateLastIterations As Long
Public P5LocateLastElement As Long
Public P5LocateLastFailure As String

' P6解析コスト診断。RCMは候補帯域幅を評価し、縮小した場合だけ
' 内部自由度番号へ適用する。帯域分解と右辺解法は分離する。
Public P6BandSolveCallCount As Long
Public P6FactorizationCount As Long
Public P6FactorizationReuseCount As Long
' V0 performance log (counters only; solver path unchanged).
Public P6PerfNewtonCount As Long
Public P6PerfIncrementCount As Long
Public P6PerfCutbackCount As Long
Public P6PerfModNewtonSkips As Long
Public P6PerfRebuildFirst As Long
Public P6PerfRebuildForce As Long
Public P6PerfRebuildForceStart As Long
Public P6PerfRebuildForceCutback As Long
Public P6PerfRebuildForceLsCut As Long
Public P6PerfRebuildForceLsOk As Long
Public P6PerfRebuildForceLsRetry As Long
Public P6PerfRebuildForceOther As Long
Public P6PerfRebuildActiveSet As Long
Public P6PerfRebuildResidual As Long
Public P6PerfRebuildStale As Long
Public P6PerfRebuildFactorInvalid As Long
Public P6SrmFs1ReuseCount As Long
Public P6PivotMin As Double
Public P6PivotMax As Double
Public P6PivotMinTrial As Double
Public P6PivotMaxTrial As Double
Public P6PivotNegCount As Long
Public P6PivotNegCountTrial As Long
Public P6PivotNearZeroCount As Long
Public P6PivotNearZeroCountTrial As Long
Public P6PivotScale As Double
Public P6LdltN As Long
Public P6LdltOk As Long
Public P6LdltFail As Long
Public P6LdltMs As Double
Public P6LdltExMax As Double
Public P6LdltErMax As Double
Public P6LdltDRatioMin As Double
Public P6LdltExMaxTrial As Double
Public P6LdltErMaxTrial As Double
Public P6LdltDRatioMinTrial As Double
Public P6PerfMarkLdltN As Long
Public P6PerfMarkLdltOk As Long
Public P6PerfMarkLdltFail As Long
Public P6PerfMarkLdltMs As Double
Public P6Ls50Tried As Long
Public P6Ls50Ok As Long
Public P6Ls50Resid As Long
Public P6Ls50Stag As Long
Public P6Ls50Worse As Long
Public P6Ls50Saved As Long
Public P6Ls50GatePolicy As Long
Public P6Ls50GateDisabled As Long
Public P6Ls50GateAlpha As Long
Public P6Ls50GateQls As Long
Public P6Ls50GateCutback As Long
Public P6Ls50GateActive As Long
Public P6Ls50GateResidual As Long
Public P6Ls50GateFactor As Long
Public P6PerfMarkLs50Tried As Long
Public P6PerfMarkLs50Ok As Long
Public P6PerfMarkLs50Resid As Long
Public P6PerfMarkLs50Stag As Long
Public P6PerfMarkLs50Worse As Long
Public P6PerfMarkLs50Saved As Long
Public P6PerfMarkLs50GatePolicy As Long
Public P6PerfMarkLs50GateDisabled As Long
Public P6PerfMarkLs50GateAlpha As Long
Public P6PerfMarkLs50GateQls As Long
Public P6PerfMarkLs50GateCutback As Long
Public P6PerfMarkLs50GateActive As Long
Public P6PerfMarkLs50GateResidual As Long
Public P6PerfMarkLs50GateFactor As Long
Public P3Ls50Pending As Boolean
Public P3Ls50Watch As Boolean
Public P3AfterCutback As Boolean
Public P6LsTries As Long
Public P6LsAccept As Long
Public P6LsA1 As Long
Public P6LsA50 As Long
Public P6LsA25 As Long
Public P6LsA12 As Long
Public P6LsOkA1 As Long
Public P6LsOkA50 As Long
Public P6LsOkA25 As Long
Public P6LsOkA12 As Long
Public P6LsAlphaSum As Double
Public P6LsAlphaMin As Double
Public P6LsAlphaMinTrial As Double
Public P6LsLastAlpha As Double
Public P6LsQCount As Long
Public P6LsQSum As Double
Public P6LsQLt50 As Long
Public P6LsQLt70 As Long
Public P6LsQLt100 As Long
Public P6LsQGe100 As Long
Public P6LsQnCount As Long
Public P6LsQnSum As Double
Public P6LsQnLt70 As Long
Public P6LsWatch As Boolean
Public P6LsWatchAfter As Double
Public P6PerfMarkLsQCount As Long
Public P6PerfMarkLsQSum As Double
Public P6PerfMarkLsQLt50 As Long
Public P6PerfMarkLsQLt70 As Long
Public P6PerfMarkLsQLt100 As Long
Public P6PerfMarkLsQGe100 As Long
Public P6PerfMarkLsQnCount As Long
Public P6PerfMarkLsQnSum As Double
Public P6PerfMarkLsQnLt70 As Long
Public P3TangentAge As Long
Public P3TangentJustRebuilt As Boolean
Public P3LastRebuildWhy As String
Public P6TanSolve0 As Long
Public P6TanSolve1 As Long
Public P6TanSolve2 As Long
Public P6TanSolve3 As Long
Public P6Age1Success As Long
Public P6Age1Resid As Long
Public P6Age1Stag As Long
Public P6Age1Active As Long
Public P6Age1Force As Long
Public P6Age1Other As Long
Public P6PerfMarkTan0 As Long
Public P6PerfMarkTan1 As Long
Public P6PerfMarkTan2 As Long
Public P6PerfMarkTan3 As Long
Public P6PerfMarkAge1Ok As Long
Public P6PerfMarkAge1Resid As Long
Public P6PerfMarkAge1Stag As Long
Public P6PerfMarkAge1Active As Long
Public P6PerfMarkAge1Force As Long
Public P6PerfMarkAge1Other As Long
' Capacity must exceed the 80-increment lookback or that slot is overwritten.
Public Const P6_ROLL_CAP As Long = 128
Public P6RollLambda(0 To 127) As Double
Public P6RollFactor(0 To 127) As Long
Public P6RollNewton(0 To 127) As Long
Public P6RollReuse(0 To 127) As Long
Public P6RollDamped(0 To 127) As Long
Public P6RollAccept(0 To 127) As Long
Public P6RollCount As Long
Public P6RollUmax(0 To 127) As Double
Public P6RollDu(0 To 127) As Double
Public P6RollPlastic(0 To 127) As Long
Public P6RollCut(0 To 127) As Long
Public P6RefC(0 To 63) As Double
Public P6RefDu(0 To 63) As Double
Public P6RefNpi(0 To 63) As Double
Public P6RefN As Long
Public P6PhysLastUmax As Double
Public P6PhysMaxDu As Double
Public P6PhysDuInit As Boolean
Public P6PhysCrossLevel As Long
Public P6PhysLastNewP As Double
Public P6PhysHaveNewP As Boolean
Public P6PhysReasonInc As Long
Public P6PhysReasons As String
Public P6LastQls As Double
Public P6LastQnext As Double
Public P6LastQlsOk As Boolean
Public P6LastQnextOk As Boolean
Public P6PhysLambda As Double
Public P6PhysStep As Double
Public P6CsvPhysHeader As Boolean
Public P6PerfMarkTimer As Double
Public P6PerfMarkAssembleMs As Double
Public P6PerfMarkEvalMs As Double
Public P6CsvSummaryHeader As Boolean
Public P6CsvTraceHeader As Boolean
Public P6CsvEventHeader As Boolean
Public P6CsvBenchHeader As Boolean
Public P6PerfMarkLsTries As Long
Public P6PerfMarkLsAccept As Long
Public P6PerfMarkLsA1 As Long
Public P6PerfMarkLsA50 As Long
Public P6PerfMarkLsA25 As Long
Public P6PerfMarkLsA12 As Long
Public P6PerfMarkLsOkA1 As Long
Public P6PerfMarkLsOkA50 As Long
Public P6PerfMarkLsOkA25 As Long
Public P6PerfMarkLsOkA12 As Long
Public P6PerfMarkLsAlphaSum As Double
Public P6PerfMarkNewton As Long
Public P6PerfMarkIncrement As Long
Public P6PerfMarkCutback As Long
Public P6PerfMarkFactor As Long
Public P6PerfMarkReuse As Long
Public P6PerfMarkBand As Long
Public P6PerfMarkModNewton As Long
Public P6PerfMarkRebuildFirst As Long
Public P6PerfMarkRebuildForce As Long
Public P6PerfMarkRebuildForceStart As Long
Public P6PerfMarkRebuildForceCutback As Long
Public P6PerfMarkRebuildForceLsCut As Long
Public P6PerfMarkRebuildForceLsOk As Long
Public P6PerfMarkRebuildForceLsRetry As Long
Public P6PerfMarkRebuildForceOther As Long
Public P6PerfMarkRebuildActiveSet As Long
Public P6PerfMarkRebuildResidual As Long
Public P6PerfMarkRebuildStale As Long
Public P6PerfMarkRebuildFactorInvalid As Long
Public P6PerfMarkTangent As Long
Public P6PerfMarkFactorMs As Double
Public P6PerfMarkSolveMs As Double
Public P6PerfMarkTangentMs As Double
Public P6PerfSummaryWritten As Boolean
Public P6FactoredBand() As Double
Public P6FactorReady As Boolean
Public P6FactorPivotTolerance As Double
Public P6BandStorageEntries As Double
Public P6BandStorageBytes As Double
Public P6FactorStorageBytes As Double
Public P6DenseStorageBytes As Double
Public P6EstimatedWorkMemoryBytes As Double
Public P6MeasuredStageTotalSeconds As Double
Public P6ModelSignature As String
Public P6MatrixStoragePolicy As String
Public P6NodeOrderingPolicy As String
Public P6RenumberingApplied As Boolean
Public P6RCMOriginalBandwidth As Long
Public P6RCMCandidateBandwidth As Long
Public P6RCMStatus As String
Public P6OptimizationPolicy As String
Public P6IntegrationStatus As String
Public P6RCMOldToNew() As Long
Public P6RCMNewToOld() As Long
Public P6RCMMapReady As Boolean
Public P6RCMCacheReady As Boolean
Public P6RCMCacheSignature As String
Public P6RCMCacheOriginalBandwidth As Long
Public P6RCMCacheCandidateBandwidth As Long
Public P6MeshInvariantReady As Boolean
Public P6MeshInvariantSignature As String
Public P6MeshInvariantOriginalBandwidth As Long
Public P6MeshInvariantCandidateBandwidth As Long
Public P6MeshInvariantRCMApplied As Boolean
Public P6MeshInvariantSymmetryError As Double
Public P6MaterialGeneration As Long
Public P6MaterialSignature As String
Public P6StrengthGeneration As Long
Public P6TangentGeneration As Long
Public P6TangentSignature As String
Public P6CSRMaterialGeneration As Long
Public P6CSRNumericGeneration As Long
Public P6CSRNumericAssemblyStatus As String
Public P6UseCSR As Boolean
Public P6CSRReady As Boolean
Public P6CSRBoundaryApplied As Boolean
Public P6CSRPatternReady As Boolean
Public P6CSRPatternSignature As String
Public P6CSRRowPtr() As Long
Public P6CSRColumnIndex() As Long
Public P6CSRDiagonalPosition() As Long
Public P6CSRElementPosition() As Long
Public P6CSROriginalValues() As Double
Public P6CSRValues() As Double
Public P6CSRNNZ As Long
Public P6CSRValueCapacity As Long
Public P6CSRStorageBytes As Double
Public P6CSRRebuildCount As Long
Public P6CSRReuseCount As Long
Public P6CSRLastAssemblyPath As String
Public P6CSRConstraintVersion As Long
Public P6CSRBoundaryConstraintVersion As Long
Public P6CSRBoundaryZeroPosition() As Long
Public P6CSRBoundaryDiagonalPosition() As Long
Public P6CSRBoundaryRHSPosition() As Long
Public P6CSRBoundaryRHSRow() As Long
Public P6CSRBoundaryRHSColumn() As Long
Public P6CSRBoundaryZeroCount As Long
Public P6CSRBoundaryDiagonalCount As Long
Public P6CSRBoundaryRHSCount As Long
Public P6CSRScatterStart() As Long
Public P6CSRScatterLocalRow() As Long
Public P6CSRScatterLocalColumn() As Long
Public P6CSRScatterPosition() As Long
Public P6CSRScatterValueIndex() As Long
Public P6CSRScatterCount As Long
Public P6CSRScatterElementCount As Long
Public P6CSRScatterTangentGeneration As Long
Public P6CSRScatterMaterialGeneration As Long
Public P6CSRScatterPatternSignature As String
Public P6CSRDiagonalZeroCount As Long
Public P6CSRFirstZeroDiagonalDof As Long
Public P6CSRMinimumDiagonal As Double
Public P6CSRDirectFallbackUsed As Boolean
Public P6BandAfterIterative As Boolean
Public P6HourglassFactor As Double
Public P6HourglassModeCount As Long
Public P6HourglassReuseCount As Long
Public P6HgShapeGeneration As Long
Public P6ProfEvalCount As Long
Public P6ProfEvalMs As Double
Public P6ProfHgMs As Double
Public P6ProfAssembleCount As Long
Public P6ProfAssembleMs As Double
Public P6ProfFactorMs As Double
Public P6ProfSolveMs As Double
Public P6ProfLoadMs As Double
Public P6ProfTangentCount As Long
Public P6ProfTangentMs As Double
Public P6ScatterRow() As Long
Public P6ScatterCol() As Long
Public P6ScatterElem() As Long
Public P6ScatterI() As Long
Public P6ScatterJ() As Long
Public P6ScatterN As Long
Public P6ScatterReady As Boolean
Public P6ScatterLastDof As Long
Public P6ScatterBand As Long
Public P6ScatterActiveGen As Long
Public P6ScatterRcm As Boolean
Public P6BandRawReady As Boolean
Public P6DeltaAssembleStreak As Long
Public P6DeltaAssembleCount As Long
Public P3ActiveSetGen As Long
Public P3NodeOnMatCache() As Boolean
Public P3LoadIndexReady As Boolean
Public P3LoadIndexActiveGen As Long
Public P3LoadIndexNodeN As Long
Public P3LoadIndexMatN As Long
Public P6FactorActivePlastic As Long
Public P6FactorFs As Double
Public P6FactorActiveGen As Long
Public P6FactorConstraintGen As Long
Public P6FactorNodeCondN As Long
Public P3ConstraintGeneration As Long
Public P3LastSuccessStepSize As Double
Public P3LineSearchSavedT() As Double
Public P3LineSearchSavedN As Long
Public P6ForceBeforeBoundaryN As Long
Public P6LastDispOutRow As Long
Public P6LastStressOutRow As Long
Public P6ForceBeforeBoundary() As Double
Public P6ForceBeforeBoundaryReady As Boolean
Public P6MixedUP As Boolean
Public P6PressureCount As Long
Public P6KrylovLast As Long
Public P6NodePressure() As Long
Public P6MixedQ() As Double
Public P6MixedCompress() As Double
Public P6MixedSchurInv() As Double
Public P6Pressure() As Double
Public P6CommittedPressure() As Double
Public P6PressureFixed() As Boolean
Public P6MixedS() As Double
Public P6MixedH() As Double
Public P6MixedC() As Double
Public P6ContinuityG() As Double
Public P6PressureResidual() As Double
Public P6QStart() As Long
Public P6QUDof() As Long
Public P6QVal() As Double
Public P6QEntryCount As Long
Public P6ConsolActive As Boolean
Public P6BiotGeometryReady As Boolean
Public P6ConsolDt As Double
Public P6ConsolTime As Double
Public P6ConsolStep As Long
Public P6MaxPressure As Double
Public P6DrainCount As Long
Public P6CSRDeviatoricReady As Boolean
Public P6MixedCoupledKrylov As Boolean
Public P6UzawaRhs() As Double
Public P6UzawaQP() As Double
Public P6UzawaRP() As Double
Public P6UzawaIterations As Long
Public P6DevFactorGeneration As Long
Public P6CSRILUValues() As Double
Public P6CSRILUReady As Boolean
Public P6SolverPolicy As String
Public P6SolverMode As String
Public P6SolverSelectionReady As Boolean
Public P6SolverReevaluationThreshold As Long
Public P6SolverEvaluationCount As Long
Public P6SolverEvaluationStatus As String
Public P6SolverMemoryLimitBytes As Double
Public P6IterativeTolerance As Double
Public P6IterativeMaxIterations As Long
Public P6GMRESRestart As Long
Public P6GMRESRestartRequested As Long
Public P6GMRESAutoAdjusted As Boolean
Public P6IterativeLastIterations As Long
Public P6IterativeLastResidual As Double
Public P6IterativeFallbackStatus As String
Public P6IterativeVectorCapacity As Long
Public P6IterativeRestartCapacity As Long
Public P6BiCGB() As Double
Public P6BiCGX() As Double
Public P6BiCGR() As Double
Public P6BiCGRHat() As Double
Public P6BiCGP() As Double
Public P6BiCGV() As Double
Public P6BiCGS() As Double
Public P6BiCGT() As Double
Public P6BiCGPHat() As Double
Public P6BiCGSHat() As Double
Public P6BiCGAX() As Double
Public P6GMRESB() As Double
Public P6GMRESX() As Double
Public P6GMRESR() As Double
Public P6GMRESZ() As Double
Public P6GMRESW() As Double
Public P6GMRESInput() As Double
Public P6GMRESBasis() As Double
Public P6GMRESH() As Double
Public P6GMRESCosine() As Double
Public P6GMRESSine() As Double
Public P6GMRESG() As Double
Public P6GMRESY() As Double
Public P6ResidualVectorCapacity As Long
Public P6CSRResidualProduct() As Double
Public P6ReactionCapacity As Long
Public P6CSRDiagonalInverse() As Double
Public P6CSRDiagonalInverseReady As Boolean
Public P6GaussCacheReady As Boolean
Public P6GaussRN(0 To 3, 0 To 15) As Double
Public P6GaussN(0 To 3, 0 To 7) As Double
Public P6InputCacheReady As Boolean
Public P6InputFingerprintReady As Boolean
Public P6InputTopologyFingerprint As Double
Public P6InputGeometryFingerprint As Double
Public P6InputMaterialFingerprint As Double
Public P6MeshRevision As Long
Public P6GeometryRevision As Long
Public P6MaterialInputRevision As Long
Public P6InputNodeCount As Long
Public P6InputMaterialCount As Long
Public P6InputElementCount As Long
Public P6InputNodeId() As Long
Public P6InputNodeX() As Double
Public P6InputNodeY() As Double
Public P6InputNodeCondX() As Long
Public P6InputNodeCondY() As Long
Public P6InputNodeDispX() As Double
Public P6InputNodeDispY() As Double
Public P6InputNodeForceX() As Double
Public P6InputNodeForceY() As Double
Public P6InputMaterialId() As Long
Public P6InputMaterialYoung() As Double
Public P6InputMaterialPoisson() As Double
Public P6InputMaterialThickness() As Double
Public P6InputMaterialWeight() As Double
Public P6InputMaterialFriction() As Double
Public P6InputMaterialCohesion() As Double
Public P6InputMaterialDilation() As Double
Public P6InputMaterialReduce() As Boolean
Public P6InputMaterialInitialActive() As Boolean
Public P6InputMaterialKind() As String
Public P6InputMaterialKs() As Double
Public P6InputMaterialAllowTension() As Boolean
Public P6InputElementId() As Long
Public P6InputElementNode() As Long
Public P6InputElementMaterial() As Long
Public P6InputHorizontalLoad As Double
Public P6InputVerticalLoad As Double
Public P6InputMode As Long
Public P6InputFSS As Double
Public P6ElasticKCache() As Double
Public P6ElasticKCacheElementCount As Long
Public P6ElasticKCacheSignature As String
Public P6ElasticKCacheReady As Boolean
Public P6ElasticKCacheStatus As String
Public P6ElasticOperatorCacheReady As Boolean
Public P6ElasticWeightCacheReady As Boolean
Public P6ElasticElementWeightCache() As Double
Public P6QuadratureArea As Double
Public P6CacheMissReason As String
Public P3CurrentElementZeroIncrement As Boolean
Public P3ActivePlasticPoint() As Boolean
Public P3ActivePlasticPointCount As Long
Public P3TrialStateInitialized As Boolean
Public femIoFileNum As Integer
Public femIoPath As String
Public femIoLastFlush As Double
Public femIoStartedAt As Double
Public femIoMode As String
Public femIoExportMode As String
Public femIoFlushSec As Double
Public femIoResumeStage As Long
Public femIoResumeFolder As String
Public femIoResumeFs As Double
Public femIoResumeFss As Double
Public femIoResumeGravity As Boolean
Public femIoResumeHasLock As Boolean
Public femIoResumePrefix As Long
Public femIoResumeIndex As Long
Public femIoResumeKind As String
Public p3StageResultLogged As Boolean
Public P3HasJointElements As Boolean
Public P3ForceRef As Double
Public P3DispRef As Double
Public P3InternalNormFree As Double
Public P3PrevResidualNorm As Double
Public P3LastTangentPlasticCount As Long
Public P3TangentStaleIters As Long
Public P3ForceTangentRebuild As Boolean
Public P3ForceRebuildReason As String
Public P3GravityReuseStatus As String
Public P3ModifiedNewtonSkips As Long
Public P3LineSearchCuts As Long
Public P3JointAssembleCount As Long
Public P6FactorJointContact() As Long
Public P6FactorJointElem() As Long
Public P6FactorJointContactN As Long
Public P6FactorYieldMask() As Byte
Public P6FactorYieldMaskN As Long
Public P6BandWorkAik() As Double
Public P6BandWorkAkj() As Double
Public P6BandWorkN As Long
Public P3LineSearchRetryCorrection As Boolean
Public P3TrialStateValid As Boolean
Public P3PlasticStateChangeCount As Long
Public P3EvalSkipCount As Long
Public P3TangentDirtyRebuildCount As Long
Public P3ElasticFastCount As Long
Public P3PolicyCacheReady As Boolean
Public P3FlowPolicyMode As Long
Public P3SrmPsiMode As Long
Public P3WeightMode As Long
Public P3StateWorkspaceElemN As Long
Public P3StateWorkspaceDofN As Long
Public P3SrmSnapWorkspaceElemN As Long
Public P3SrmSnapWorkspaceDofN As Long

Private Sub P6InvalidateSettingCache()
    ' 設定値の本体はCONTROL.FEMReadSetting。ここでは解法選択だけ落とす。
    P6SolverSelectionReady = False
    P6SolverReevaluationThreshold = 0
    P6SolverEvaluationStatus = "invalidated"
    P3PolicyCacheReady = False
End Sub

Public Function P6FingerprintStep(ByVal currentValue As Double, ByVal nextValue As Double) As Double
    Dim updatedValue As Double
    updatedValue = currentValue * 131# + nextValue
    updatedValue = updatedValue - Fix(updatedValue / 2147483629#) * 2147483629#
    If updatedValue < 0# Then updatedValue = updatedValue + 2147483629#
    P6FingerprintStep = updatedValue
End Function

Public Function P6QuantizeScaled(ByVal value As Double, ByVal xscale As Double) As Double
    Dim scaledValue As Double
    scaledValue = value * xscale
    If scaledValue >= 0# Then
        P6QuantizeScaled = Int(scaledValue + 0.5)
    Else
        P6QuantizeScaled = -Int(-scaledValue + 0.5)
    End If
End Function

Public Function P6NextRevision(ByVal currentRevision As Long) As Long
    If currentRevision >= 2147483647 Then
        P6NextRevision = 1
    Else
        P6NextRevision = currentRevision + 1
    End If
End Function

Public Function P6ReadSetting(ByVal keyName As String, ByVal defaultValue As Double) As Double
    P6ReadSetting = FEMReadSetting(keyName, defaultValue)
End Function

Public Function P6ReadTextSetting(ByVal keyName As String, ByVal defaultValue As String) As String
    P6ReadTextSetting = FEMReadTextSetting(keyName, defaultValue)
End Function

Public Function P6ElasticWeightCacheUsable() As Boolean
    P6ElasticWeightCacheUsable = False
    If Not P6ElasticWeightCacheReady Then Exit Function
    If P6ElasticKCacheElementCount < 1 Then Exit Function
    On Error GoTo Fail
    If LBound(P6ElasticElementWeightCache) <> 0 Then GoTo Fail
    If UBound(P6ElasticElementWeightCache) <> P6ElasticKCacheElementCount - 1 Then GoTo Fail
    P6ElasticWeightCacheUsable = True
    Exit Function
Fail:
    P6ElasticWeightCacheReady = False
    P6ElasticWeightCacheUsable = False
End Function

Public Sub ResetAnalysisState()
  Dim resultId As Long
  AnalysisOK = False
  ResultStatus = RESULT_NOT_RUN
  AnalysisMessage = ""
  AnalysisErrorNumber = 0
  P3UserCancel = False
  P3SearchFatal = False
  FailureElement = -1
  FailureGaussPoint = -1
  FailureIncrement = -1
  FailureIteration = -1
  MatrixFactored = False
  CurrentIncrement = -1
  CurrentIteration = -1
  ResidualNormFull = 0#
  ResidualNormFree = 0#
  RelativeResidualFree = 0#
  ForceNormFree = 0#
  MaxAbsResidualFree = 0#
  EnergyError = 0#
  MatrixSymmetryError = 0#
  TotalAppliedLoadX = 0#
  TotalAppliedLoadY = 0#
  TotalSelfWeight = 0#
  P6WeightMode = "TOTAL"
  P6WeightNote = "全応力γ。Q8 2×2（剛性・内力と同じ積分）。水位・浮力なし"
  MinDetJCorner = 1E+308
  MaxDetJCorner = -1E+308
  MinDetJGauss = 1E+308
  MaxDetJGauss = -1E+308
  TimeInput = 0#
  TimeElementStiffness = 0#
  TimeAssembly = 0#
  TimeSolve = 0#
  TimePlastic = 0#
  TimeOutput = 0#
  P6HourglassReuseCount = 0
  P6ProfEvalCount = 0
  P6ProfEvalMs = 0#
  P6ProfHgMs = 0#
  P6ProfAssembleCount = 0
  P6ProfAssembleMs = 0#
  P6ProfFactorMs = 0#
  P6ProfSolveMs = 0#
  P6ProfLoadMs = 0#
  P6ProfTangentCount = 0
  P6ProfTangentMs = 0#
  P6ScatterReady = False
  P3LoadIndexReady = False
  P6FactorNodeCondN = -1
  P6FactorConstraintGen = -1
  P3ConstraintGeneration = 1
  P3LineSearchSavedN = -1
  P6ForceBeforeBoundaryN = -1
  P3LastSuccessStepSize = 0#
  P3PolicyCacheReady = False
  P3StateWorkspaceElemN = -1
  P3StateWorkspaceDofN = -1
  P3SrmSnapWorkspaceElemN = -1
  P3SrmSnapWorkspaceDofN = -1
  P6FactorReady = False
  P3UseNonlinearResidual = False
  P3CurrentStrengthFactor = 1#
  P3LastConvergedLoadFactor = 0#
  P3LastConvergedIncrement = 0
  P3SuccessfulIncrementCount = 0
  P3GlobalIterationCount = 0
  P3PlasticPointUpdateCount = 0
  P3RetryCount = 0
  P3DefaultIncrementCount = P3_DEFAULT_INCREMENT_COUNT
  P3MaxTrialDisp = 0#
  P3MaxCorrection = 0#
  P3LastRelativeResidual = 0#
  P3MaxBoundaryDispError = 0#
  P3RelativeBoundaryDispError = 0#
  P6ReactionSumX = 0#
  P6ReactionSumY = 0#
  P5ModelScale = 0#
  P5GeometryTolerance = 0#
  P5AreaTolerance = 0#
  P5MinElementArea = 0#
  P5MinJacobian = 0#
  P5MaxAspectRatio = 0#
  P5MinAngleDeg = 0#
  P5MaxAngleDeg = 0#
  P5BoundaryEdgeCount = 0
  P5SharedEdgeCount = 0
  P5MaterialInterfaceCount = 0
  P5IsolatedNodeCount = 0
  P5ConnectedComponentCount = 0
  P5BoundarySelfIntersectionCount = 0
  P5LocateCallCount = 0
  P5LocateTotalIterations = 0
  P5LocateMaxIterations = 0
  P5LocateLastIterations = 0
  P5LocateLastElement = 0
  P5LocateLastFailure = vbNullString
  P6BandSolveCallCount = 0
  P6FactorizationCount = 0
  P6FactorizationReuseCount = 0
  P6PerfNewtonCount = 0
  P6PerfIncrementCount = 0
  P6PerfCutbackCount = 0
  P6PerfModNewtonSkips = 0
  P6PerfRebuildFirst = 0
  P6PerfRebuildForce = 0
  P6PerfRebuildForceStart = 0
  P6PerfRebuildForceCutback = 0
  P6PerfRebuildForceLsCut = 0
  P6PerfRebuildForceLsOk = 0
  P6PerfRebuildForceLsRetry = 0
  P6PerfRebuildForceOther = 0
  P6PerfRebuildActiveSet = 0
  P6PerfRebuildResidual = 0
  P6PerfRebuildStale = 0
  P6PerfRebuildFactorInvalid = 0
  P6SrmFs1ReuseCount = 0
  P6PivotMin = 1E+308
  P6PivotMinTrial = 1E+308
  P6PivotNegCount = 0
  P6PivotNegCountTrial = 0
  P6PivotNearZeroCount = 0
  P6PivotNearZeroCountTrial = 0
  P6PivotScale = 1#
  P6LsTries = 0
  P6LsAccept = 0
  P6LsA1 = 0
  P6LsA50 = 0
  P6LsA25 = 0
  P6LsA12 = 0
  P6LsOkA1 = 0
  P6LsOkA50 = 0
  P6LsOkA25 = 0
  P6LsOkA12 = 0
  P6LsAlphaSum = 0#
  P6LsAlphaMin = 1E+308
  P6LsAlphaMinTrial = 1E+308
  P6LsLastAlpha = 0#
  P6LsQCount = 0
  P6LsQSum = 0#
  P6LsQLt50 = 0
  P6LsQLt70 = 0
  P6LsQLt100 = 0
  P6LsQGe100 = 0
  P6LsQnCount = 0
  P6LsQnSum = 0#
  P6LsQnLt70 = 0
  P6LsWatch = False
  P6LsWatchAfter = 0#
  P6Ls50Tried = 0
  P6Ls50Ok = 0
  P6Ls50Resid = 0
  P6Ls50Stag = 0
  P6Ls50Worse = 0
  P6Ls50Saved = 0
  P6Ls50GatePolicy = 0
  P6Ls50GateDisabled = 0
  P6Ls50GateAlpha = 0
  P6Ls50GateQls = 0
  P6Ls50GateCutback = 0
  P6Ls50GateActive = 0
  P6Ls50GateResidual = 0
  P6Ls50GateFactor = 0
  P3Ls50Pending = False
  P3Ls50Watch = False
  P3Ls50Enabled = False
  P6BandBenchOnly = False
  P3AfterCutback = False
  P3TangentAge = 0
  P3TangentJustRebuilt = False
  P3LastRebuildWhy = vbNullString
  P6TanSolve0 = 0
  P6TanSolve1 = 0
  P6TanSolve2 = 0
  P6TanSolve3 = 0
  P6Age1Success = 0
  P6Age1Resid = 0
  P6Age1Stag = 0
  P6Age1Active = 0
  P6Age1Force = 0
  P6Age1Other = 0
  P6RollCount = 0
  P6PhysResetTrial
  P6PivotMax = 0#
  P6PivotMaxTrial = 0#
  P6LdltN = 0
  P6LdltOk = 0
  P6LdltFail = 0
  P6LdltMs = 0#
  P6LdltExMax = 0#
  P6LdltErMax = 0#
  P6LdltDRatioMin = 1E+308
  P6LdltExMaxTrial = 0#
  P6LdltErMaxTrial = 0#
  P6LdltDRatioMinTrial = 1E+308
  P6CsvSummaryHeader = False
  P6CsvTraceHeader = False
  P6CsvEventHeader = False
  P6CsvPhysHeader = False
  P6CsvBenchHeader = False
  P6ResetBandBench
  P3ForceRebuildReason = vbNullString
  P3GravityReuseStatus = "NONE"
  P6PerfSummaryWritten = False
  P6PerfMark
  P6FactorReady = False
  P6FactorPivotTolerance = 0#
  Erase P6FactoredBand
  P6BandStorageEntries = 0#
  P6BandStorageBytes = 0#
  P6FactorStorageBytes = 0#
  P6DenseStorageBytes = 0#
  P6EstimatedWorkMemoryBytes = 0#
  P6MeasuredStageTotalSeconds = 0#
  P6ModelSignature = vbNullString
  P6MatrixStoragePolicy = vbNullString
  P6NodeOrderingPolicy = vbNullString
  P6RenumberingApplied = False
  P6RCMOriginalBandwidth = 0
  P6RCMCandidateBandwidth = 0
  P6RCMStatus = vbNullString
  P6OptimizationPolicy = vbNullString
  P6IntegrationStatus = vbNullString
  P6RCMMapReady = False
  ' RCM写像とCSR構造はメッシュが変わるまで保持する。
  ' 解析ごとに変わるのは数値行列、境界条件、右辺だけである。
  P6UseCSR = False
  P6CSRReady = False
  P6CSRBoundaryApplied = False
  P6CSRConstraintVersion = 0
  P6CSRBoundaryConstraintVersion = -1
  P6CSRBoundaryZeroCount = 0
  P6CSRBoundaryDiagonalCount = 0
  P6CSRBoundaryRHSCount = 0
  P6CSRRebuildCount = 0
  P6CSRReuseCount = 0
  P6CSRLastAssemblyPath = vbNullString
  P6CSRScatterCount = 0
  P6CSRScatterElementCount = -1
  P6CSRScatterTangentGeneration = -1
  P6CSRScatterMaterialGeneration = -1
  P6CSRScatterPatternSignature = vbNullString
  P6CSRDiagonalZeroCount = 0
  P6CSRFirstZeroDiagonalDof = -1
  P6CSRMinimumDiagonal = 0#
  P6CSRDirectFallbackUsed = False
  P6BandAfterIterative = False
  P6HourglassFactor = P6_HOURGLASS_FACTOR_DEFAULT
  P6HourglassModeCount = 0
  P6ForceBeforeBoundaryReady = False
  Erase P6ForceBeforeBoundary
  P6MixedUP = False
  P6PressureCount = 0
  P6KrylovLast = -1
  P6CSRDeviatoricReady = False
  P6MixedCoupledKrylov = False
  P6UzawaIterations = 0
  P6DevFactorGeneration = -1
  P6CSRILUReady = False
  P6ConsolActive = False
  P6BiotGeometryReady = False
  P6ConsolDt = 0#
  P6ConsolTime = 0#
  P6ConsolStep = 0
  P6MaxPressure = 0#
  P6DrainCount = 0
  P6QEntryCount = 0
  Erase P6NodePressure
  Erase P6MixedQ
  Erase P6MixedCompress
  Erase P6MixedSchurInv
  Erase P6Pressure
  Erase P6CommittedPressure
  Erase P6PressureFixed
  Erase P6MixedS
  Erase P6MixedH
  Erase P6MixedC
  Erase P6ContinuityG
  Erase P6PressureResidual
  Erase P6QStart
  Erase P6QUDof
  Erase P6QVal
  Erase P6UzawaRhs
  Erase P6UzawaQP
  Erase P6UzawaRP
  Erase P6CSRILUValues
  Erase P6CSRScatterStart
  Erase P6CSRScatterLocalRow
  Erase P6CSRScatterLocalColumn
  Erase P6CSRScatterPosition
  Erase P6CSRScatterValueIndex
  Erase P6CSRBoundaryZeroPosition
  Erase P6CSRBoundaryDiagonalPosition
  Erase P6CSRBoundaryRHSPosition
  Erase P6CSRBoundaryRHSRow
  Erase P6CSRBoundaryRHSColumn
  ' CSR数値配列と対角前処理配列は同一メッシュで再利用する。
  ' メッシュ変更時はP6PrepareMeshCachesで明示的に破棄する。
  P6CSRDiagonalInverseReady = False
  P6CSRNumericAssemblyStatus = vbNullString
  P6InvalidateSettingCache
  P6SolverPolicy = vbNullString
  P6SolverMode = vbNullString
  P6SolverSelectionReady = False
  P6SolverReevaluationThreshold = 0
  P6SolverEvaluationCount = 0
  P6SolverEvaluationStatus = vbNullString
  P6SolverMemoryLimitBytes = 0#
  P6ElasticOperatorCacheReady = False
  P6QuadratureArea = 0#
  P6CacheMissReason = vbNullString
  P3CurrentElementZeroIncrement = False
  P3TrialStateInitialized = False
  P3ActivePlasticPointCount = 0
  Erase P3ActivePlasticPoint
  P6IterativeTolerance = 0#
  P6IterativeMaxIterations = 0
  P6GMRESRestart = 0
  P6IterativeLastIterations = 0
  P6IterativeLastResidual = 0#
  P6IterativeFallbackStatus = vbNullString
  ModelRevision = ModelRevision + 1
  ResultRevision = 0
  P1Formulation = P1_FORMULATION_PLANE_STRAIN
  P1ResultReady = False
  P1ResultModelRevision = 0
  For resultId = 0 To P1_RESULT_COUNT - 1
    P1ResultAvailable(resultId) = False
  Next resultId
  NumberOfFreeNode = 0
  NumberOfNode = 0
  NumberOfMaterial = 0
  NumberOfElement = 0
  BandWidth = 0
  FEMResetRunLog
End Sub

Public Sub SetP0SilentMode(ByVal isSilent As Boolean)
  SuppressUserMessages = isSilent
End Sub

Public Sub FEMInvalidateP6SettingCache()
  P6InvalidateSettingCache
End Sub

Public Sub SetAnalysisFailure(ByVal statusText As String, ByVal messageText As String, ByVal errorNumber As Long, _
                              ByVal elementId As Long, ByVal gaussPointId As Long, _
                              ByVal incrementId As Long, ByVal iterationId As Long)
  AnalysisOK = False
  ResultStatus = statusText
  AnalysisMessage = messageText
  AnalysisErrorNumber = errorNumber
  P1ResultReady = False
  FailureElement = elementId
  FailureGaussPoint = gaussPointId
  FailureIncrement = incrementId
  FailureIteration = iterationId
  ResultRevision = 0
  If errorNumber = 18 Then
    P3UserCancel = True
    P3SearchFatal = True
    statusText = RESULT_NONCONVERGED
    If Len(messageText) = 0 Then messageText = "ユーザーが解析を中断しました。"
    ResultStatus = statusText
    AnalysisMessage = messageText
  ElseIf statusText = RESULT_INPUT_ERROR Or statusText = RESULT_CAPACITY_ERROR Or statusText = RESULT_RUNTIME_ERROR Then
    P3SearchFatal = True
  ElseIf statusText = RESULT_MATERIAL_ERROR And (errorNumber = vbObjectError + 3240 Or InStr(1, messageText, "非対称", vbTextCompare) > 0) Then
    P3SearchFatal = True
  End If
  If P3SrmTrialRunning Then
    If Len(P3SrmTrialFailNote) = 0 Then
      P3SrmTrialFailNote = statusText & " " & messageText
    ElseIf statusText = RESULT_NONCONVERGED And InStr(1, P3SrmTrialFailNote, RESULT_NONCONVERGED, vbTextCompare) = 0 Then
      P3SrmTrialFailNote = statusText & " " & messageText
    End If
  Else
    On Error Resume Next
    FEMAppendRunLog FEMLastProc, statusText, messageText
    On Error GoTo 0
  End If
End Sub

Private Function P6Elapsed(ByVal startedAt As Double) As Double
  P6Elapsed = Timer - startedAt
  If P6Elapsed < 0# Then P6Elapsed = P6Elapsed + 86400#
End Function

Public Function P6ElapsedMs(ByVal startedAt As Double) As Double
  P6ElapsedMs = P6Elapsed(startedAt) * 1000#
End Function

Public Sub P6InvalidateSpeedCaches()
  P6ScatterReady = False
  P6BandRawReady = False
  P6DeltaAssembleStreak = 0
  P3LoadIndexReady = False
  P6FactorNodeCondN = -1
  P6FactorConstraintGen = -1
  P6FactorReady = False
End Sub

Public Sub P6InvalidateActiveDependentCaches()
  P6InvalidateSpeedCaches
  P3ActiveSetGen = P3ActiveSetGen + 1
  If P3ActiveSetGen <= 0 Then P3ActiveSetGen = 1
End Sub













# 土の弾塑性FEM：VBAのみで行う高速化計画

日付：2026-10-04

## 1. 実行環境と変更範囲

外部DLL、外部実行ファイル、外部数値ソルバーへの解析委譲を使わない。Excel/VBA内で解析を完結させる。Pythonは本調査で添付ZIP・VBAソース・ログを読むために用いたものであり、提案する解析の実行依存には含めない。

既存のBand LUを基準ソルバーとして残す。まず分解回数、荷重増分数、高価なSRM試行数を減らす。収束許容値、材料強度、流れ則、メッシュ、荷重履歴、SRMの終了精度を緩めて速度を得る案は対象外とする。

対象は `2DSoilFEM_202609260838_out.zip` 内の `2DSoilFEM_202609260838.xlsm` と計測ログ。VBAソースを今回再抽出して静的確認した。VBA実行・新アルゴリズムの実装・再解析は行っていない。下記は検証計画であり、実測した高速化結果ではない。

## 2. 再確認したボトルネック

原ログ `perf_summary.txt` より、総時間4154.676秒、行列分解3094.531秒、分解8016回、再利用4408回。分解は総時間の74.483%を占める。

接線更新理由のカウンターは次のとおり。分類合計は8016で分解回数と一致するが、それぞれを独立に除去できるという意味ではない。

| 更新理由 | 回数 | 更新理由全体の比率 |
|---|---:|---:|
| FIRST | 3057 | 38.14% |
| RESIDUAL | 2450 | 30.56% |
| FORCE | 2103 | 26.24% |
| STAGNATION | 353 | 4.40% |
| ACTIVESET | 53 | 0.66% |

FORCEの内訳はSTAGE_START=17、CUTBACK=15、LINESEARCH_CUT=0、LSOK分類=2016、LSRETRY=55。注意：`InvalidateTangent` のLSOKカウンターは内側Select CaseのElseで増えるため、2016回すべてを文字どおりのLINESEARCH_OK要求と解釈してはいけない。LS50などの他の強制更新理由もこの集約へ入る。次回計測では理由文字列別に分離する。

RCM再番号付け、接線再利用、ラインサーチ後の条件付き再利用、収束判定を行列分解より先に実施する処理、trial-stateキャッシュは既に存在する。これらを「未実装の新機能」として再提案しない。

## 3. 推奨する実装・検証順序

| 段階 | 内容 | 主な修正対象 | 位置付け |
|---|---|---|---|
| V0 | 基準保存、更新理由と増分履歴の計測追加 | 性能ログ、更新理由カウンター | 全変更に先行 |
| V1 | 同一ステージ内の変位予測子 | P3RunLoadStagesに独立処理追加 | 最初の小改修 |
| V2a | 塑性状態での増分間接線・LU再利用を限定試験 | P3ShouldRebuildTangent等 | 分解回数への主要対策 |
| V2b | 残差改善率・実測費用による接線更新判定 | 同上 | V2aとは別フラグ |
| V3 | 費用を考慮した保護付きFs探索 | NextFsByBracket、P3RunStrengthReduction | 固定Fsソルバー試験と分離 |
| V4 | 深さ1の保護付きAnderson加速 | 反復補正、履歴、残差評価 | 凍結接線区間のみ |
| V5 | 現行LU因子を前処理とするVBA製GMRES | 行列・因子管理、行列ベクトル積、線形求解 | 中規模改修 |
| 条件付き | 全サブステップ整合接線 | P2MaterialPointUpdate | 実際の複数分割使用頻度で繰上げ |

## 4. V1：変位予測子

現状は `P3CopyCommittedDispToTrial` で確定変位をそのまま初期値にする。同一Fs、同一ステージ、同じ有効要素・拘束条件の下で、直前の確定変位増分から自由DOFの初期推定を作る。

    u_pred = u_n + beta * (delta_lambda_new / delta_lambda_old) * (u_n - u_(n-1))

beta=1を基本候補とし、外挿量制限・前増分の品質に応じた制限を設ける。数値の具体的な上限は検証用パラメータとし、検証なしに固定しない。

### 実装上の必須条件

- 変更するのはtrial変位だけ。確定応力、塑性ひずみ、その他の内部変数、および `incrementBase` は変更しない。
- `P3CopyCommittedDispToTrial` 自体を外挿処理に置き換えない。この関数はロールバックにも使われる。増分開始専用の `P3ApplyIncrementPredictor` を新設する案とする。
- 拘束DOFは外挿しない。現在の目標荷重係数に対応する所定変位を与える。
- trial変位を変更したら `P3TrialStateValid` 等の対応するキャッシュを無効化し、最新変位で内力・応力を再評価する。
- ステージ・Fs・拘束条件・有効要素変更時、cutback時、必要な履歴不足時には従来初期値を使う。
- 予測値による更新が失敗したら、確定状態から従来初期値でやり直す。予測子単独の失敗をFsのFAILへ直結させない。

現行の刻み制御では、2反復以下で1.5倍、3〜5反復で維持、6〜8反復で0.7倍、それ以上で0.5倍になる。まずこの規則を固定したまま予測子だけを比較する。反復数が減ることで刻み増大へつながる可能性はあるが、実測が必要である。刻み制御そのものの変更は別フラグで、荷重―変位曲線と刻み依存性を検証する。

## 5. V2：適応的な接線更新とLU再利用

`P3ShouldRebuildTangent` は既に修正Newtonであり、残差改善率、塑性点変化、接線の使用回数、LS50状態を監視している。新しい方針は、単なる再利用機能の追加ではなく、既存ルールの限定的な改良である。

### V2a：増分間再利用

現行では増分初回に塑性点が1点でも存在するとFIRST更新の対象になる。微小な増分が連続する区間でも発生するため、ここを最初の検証対象にする。

同一Fs・ステージ・境界条件・有効要素で、直前増分が正常収束し、cutback直後でなく、接触状態・材料分岐・接線変化が小さい候補だけを選ぶ。まずは次増分の最初の補正1回だけに旧接線を用い、残差改善が不十分なら即座に再構築する方式から始める。

古い接線を最新接線の正確な因子と偽って扱わない。これは修正Newtonの近似Jacobianとしての使用であり、最新の内力と残差を用いて非線形平衡を確認する。現行の因子有効性署名が厳密な同一行列再利用用であるなら、それを破壊せず、近似Jacobianとしての再利用資格を別に管理する。

残差改善率は同じ増分・同じ目標荷重で補正前後を比較する。直前増分のほぼゼロの残差と、新増分の残差を比較しない。

### V2b：反復内の更新選別

通常分岐では残差が前回の0.7倍より大きいとRESID更新、再利用カウンターが3に達するとSTAG更新する。これらを無条件に緩めるのではなく、残差改善率の推移、材料分岐・接触変化、接線差の指標、更新後の改善実績、因子分解と追加反復の実測費用を使う。

残差から必要反復数を予測する方法は補助的なヒューリスティックにとどめる。破壊近傍では収束率が変わるため予測だけで更新を省略しない。最大再利用数、停滞時即更新、更新後再試行、従来経路へのフォールバックを残す。

塑性点の総数だけでは、MCの面・稜線・頂点の切替や同数の別の点への入れ替わりを検出できない。現在のXOR監視を維持した上で、必要な材料分岐変化の診断を追加する。

## 6. V3：SRM探索の総費用を減らす

最新の適応探索研究を参考に、失敗側・収束側の実測費用を含めて次のFsを選ぶ。現行の二分法を必ずフォールバックとして残す。探索点を端点へ過度に寄せず、区間が縮まらないときは二分へ戻す。

補助指標を使う場合も、PASS/FAILの±1をそのまま連続根探索関数と見なさない。収束判定そのもの、メッシュ、材料・初期条件を変更しない。時間切れや高速化経路の失敗は未判定として基準経路へ戻し、FAIL端点の根拠にしない。

現在の基準区間はPASS=2.2625、FAIL=2.275、幅0.0125。設定SRM_TOL=0.02だけでなく、同等精度比較では実際の最終区間幅も合わせる。この区間は数値判定に基づくもので、保証された極限解析の上下界ではない。

## 7. V4：深さ1の保護付きAnderson加速

凍結接線を使う修正Newton区間に限定し、前回と今回の変位・補正量の差から加速候補を作る。追加計算は主としてベクトルの内積と線形結合であり、VBA単独実装の候補になる。

接線更新、新増分、Fs変更、cutback、材料・接触の大きな分岐変化で履歴をリセットする。小さい分母、過大混合係数、非有限値、材料更新失敗を棄却する。候補の実残差を評価し、悪化時は履歴を捨てて既存Newton/ラインサーチ経路へ戻す。

参照するnumgeo実装ではAndersonとラインサーチの同時使用は非対応である。これは両者が理論上いかなる構成でも組み合わせられないという主張ではない。現行LS50に機械的に重ねず、未減衰補正を前提とする独立分岐として実装・比較する。

## 8. V5：VBA製GMRES＋旧LU前処理

最新の接線方程式 `K_k * du = -r_k` を解き、以前の接線のLU因子を前処理として使う。逆行列を形成しない。行列ベクトル積、前後進代入、Arnoldi直交化、小さいHessenberg系の求解をVBA内で行う。

最初の試験では最新接線の帯行列を使い、CSRへの全面移行を必須にしない。最新接線行列と前処理因子は別管理し、同じ配列への破壊的上書きを避ける。境界処理・DOF順序を一致させる。

最初は線形求解精度も緩めず、真の残差 `b - K_k * du` を同じ基準で確認する。前処理付き残差だけで完了としない。1回の通常GMRES求解の途中では前処理を固定し、更新は求解の再開始時に行う。途中変更を許す場合はFGMRES等の別設計が必要になるため、初版には含めない。

Krylov反復が増えた場合は前処理更新、それでも不調なら最新接線の既存Band LUへ戻す。GMRESの失敗をSRMのFAILへ直結させない。

現行コードにはGMRES等が存在する一方、本番経路は `P6SelectSolverMode` でBAND_LUに固定される。設定変更だけで新方式が有効になるわけではない。最新行列と旧因子の分離、残差検査、既存フォールバックを実装する必要がある。

## 9. 条件付き対策：全サブステップの整合接線

`P2MaterialPointUpdate` は最後のサブステップの接線を返す。全ひずみ増分から最終応力までの微分と整合しない可能性があるため、複数サブステップの呼出頻度と方向微分誤差を先に計測する。

使用が多く誤差も大きければV4/V5より先に検証する。MCの面・稜線・頂点を含む分岐を扱い、確定状態を固定した微分検証を行う。2026年Buiらの対象はMCC/CASMであり、MCへ式をそのまま移植するものではない。

## 10. 比較試験と採否

最初は固定Fs=2.20と2.25で、基準、V1のみ、V2aのみを比較する。通過後に2.2625、および基準で非収束の2.275・2.30へ進む。V2b、V3、V4、V5は単独評価の後で組み合わせる。

記録するもの：総時間、分解回数、再利用回数、更新理由の正確な文字列、Newton反復、増分数、増分幅、cutback、ラインサーチ評価回数、サブステップ分布、候補棄却、フォールバック、線形・非線形残差、変位、塑性域、材料収束、力学監視。

線形弾性、塑性化が進むケース、破壊近傍、c=0、異なる流れ則の対応範囲、接触・拘束変更、ステージ切替も回帰試験に含める。対象機能の範囲を明示し、非対応の設定を新方式で対応済みと扱わない。

旧FAIL点が改良法で収束した場合は即座に不合格とせず、残差・材料条件・過大変形・刻み依存性・機構を検証する。旧非収束を数学的な破壊証明としない。

### 算術的な感度

分解1回の単価と他の処理費用が変わらない仮定の下で、分解回数25%削減なら全体約1.23倍、50%削減なら約1.59倍。これは実測した高速化率でも予測値でもない。追加反復・候補残差評価・GMRESの費用を差し引いて総時間で判断する。

## 11. 参照研究と本計画の区別

- Karátson, Sysala, Béreš (2025), *Quasi-Newton iterative solution approaches for nonsmooth elliptic operators with applications to elasto-plasticity*, Computers & Mathematics with Applications 178, 61–80. DOI: `10.1016/j.camwa.2024.11.022`. 非滑らかな弾塑性に対するquasi-Newton/variable preconditioningを参照。V2の具体的な閾値・フォールバックは本コード向けの独自検証案であり、論文の完全再現ではない。
- Sysala et al. (2025), *Advanced continuation and iterative methods for slope stability analysis in 3D*, Computers & Structures 315, 107842. DOI: `10.1016/j.compstruc.2025.107842`. 間接継続、inexact Newton-like、前処理付きdeflated Krylovの組合せを参照。V5の旧Band LU前処理GMRESは本コード向けの設計案であり、論文全体の直接移植ではない。
- Sun, Chen (2026), *Strength reduction search algorithms for slope stability analysis using spectral finite element method: From ‘brittle’ to ‘ductile’ failure*, Computers and Geotechnics 196, 108175. DOI: `10.1016/j.compgeo.2026.108175`. Fs探索費用と破壊形態の関係を参照。論文の削減率を本VBAの期待倍率として使わない。
- Bui, Niníc, Meschke (2026), *Implicit sub-stepping scheme for critical state soil models*, Engineering with Computers 42, article 76. DOI: `10.1007/s00366-026-02309-1`. 全サブステップを通した整合接線を参照。対象構成式の相違を考慮する。
- numgeo公式文書、Theory / Solution method / Accelerators。深さ1・2のAnderson加速と保護条件、参照実装でのラインサーチとの排他的使用を参照。

## 12. 原ログ・コード抜粋

以下は今回の再抽出で確認した原文。行番号は改行正規化後のUTF-8抽出ソースに対応する。後半のコードは提案の実装ではなく、現行コードの引用である。

### 原ログ perf_summary.txt

```text
========== PERFORMANCE SUMMARY ==========
Ver=202609260838
ElapsedTotal=4154.676 sec
SRMTrialCount=17
Fs1ReuseCount=1
FOS_PASS=2.263
FOS_FAIL=2.275
FOS_MID=2.269
FOS_WIDTH=0.013
FOS=2.269 ± 0.006
NewtonIterationCount=12424
LoadIncrementCount=3159
CutbackCount=17
FactorizationCount=8016
FactorizationReuseCount=4408
AVG_FACTOR_REUSE=0.550
BandSolveCallCount=12424
TangentBuildCount=8017
TangentReuseCount=4408
REBUILD_FIRST_ITER=3057
REBUILD_FORCE=2103
REBUILD_FORCE_START=17
REBUILD_FORCE_CUTBACK=15
REBUILD_FORCE_LSCUT=0
REBUILD_FORCE_LSOK=2016
REBUILD_FORCE_LSRETRY=55
REBUILD_ACTIVESET=53
REBUILD_RESIDUAL=2450
REBUILD_STAGNATION=353
REBUILD_FACTOR_INVALID=0
MaxDisp=1.434E+00
RelativeResidualFree=4.973E-06
EnergyError=2.640E-07
MatrixSymmetryError=1.661E-16
MinPivot=2.001E+03
NegativePivotCount=0
NearZeroPivotCount=0
PivotMax=7.913E+04
PivotRatio=2.529E-02
LsTries=16680
LsAccept=12352
LsA1=9169
LsA50=2402
LsA25=705
LsA12=76
LsOkA1=0
LsOkA50=2402
LsOkA25=705
LsOkA12=76
LsAlphaMean=0.855
LsAlphaMin=0.125
PlasticPoints=625/1224
TimeFactorization_ms=3094531
TimeLinearSolve_ms=157061
TimeTangent_ms=249414
TimeAssembly_ms=25754
TimeEval_ms=581342
SolverMode=BAND_LU
BandLUKernel=NEW
FactorAvgMs=386.044
LDLTShadowCount=0
LDLTOkCount=0
LDLTFailCount=0
LDLTFactorMs=0.000
LimitStateMonitor=SRM_SHADOW
LimitStateNote=per-Fs mechanical_status is in perf_summary.csv

=========================================

```

### FEMEngine.bas:1317–1350

```vb
 1317: Private Sub InvalidateTangent(ByVal reason As String)
 1318:   P6FactorReady = False
 1319:   P3ForceTangentRebuild = True
 1320:   P3ForceRebuildReason = reason
 1321:   P3Ls50Pending = False
 1322:   P3Ls50Watch = False
 1323:   P3LastRebuildWhy = reason
 1324:   Select Case reason
 1325:     Case "FIRST"
 1326:       P6PerfRebuildFirst = P6PerfRebuildFirst + 1
 1327:     Case "RESID"
 1328:       P6PerfRebuildResidual = P6PerfRebuildResidual + 1
 1329:     Case "STAG", "LS50_STAGNATION"
 1330:       P6PerfRebuildStale = P6PerfRebuildStale + 1
 1331:     Case "ACTIVE"
 1332:       P6PerfRebuildActiveSet = P6PerfRebuildActiveSet + 1
 1333:     Case "FACINV"
 1334:       P6PerfRebuildFactorInvalid = P6PerfRebuildFactorInvalid + 1
 1335:     Case Else
 1336:       P6PerfRebuildForce = P6PerfRebuildForce + 1
 1337:       Select Case reason
 1338:         Case "STAGE_START"
 1339:           P6PerfRebuildForceStart = P6PerfRebuildForceStart + 1
 1340:         Case "CUTBACK"
 1341:           P6PerfRebuildForceCutback = P6PerfRebuildForceCutback + 1
 1342:         Case "LINESEARCH_CUT"
 1343:           P6PerfRebuildForceLsCut = P6PerfRebuildForceLsCut + 1
 1344:         Case "LINESEARCH_RETRY"
 1345:           P6PerfRebuildForceLsRetry = P6PerfRebuildForceLsRetry + 1
 1346:         Case Else
 1347:           P6PerfRebuildForceLsOk = P6PerfRebuildForceLsOk + 1
 1348:       End Select
 1349:   End Select
 1350: End Sub
```

### FEMEngine.bas:1352–1382

```vb
 1352: Private Function P3ShouldRebuildTangent(ByVal localIteration As Long, ByVal residualNow As Double) As Boolean
 1353:   Dim plasticJump As Long, plasticRef As Long, xorLimit As Long, factorXor As Long
 1354:   Dim ls50 As Boolean, watch As Boolean, heldReason As String
 1355:   P3ShouldRebuildTangent = True
 1356:   ls50 = P3Ls50Pending
 1357:   watch = P3Ls50Watch
 1358:   P3Ls50Pending = False
 1359:   P3Ls50Watch = False
 1360:   If P3ForceTangentRebuild Then
 1361:     heldReason = P3ForceRebuildReason
 1362:     P3ForceTangentRebuild = False
 1363:     P3ForceRebuildReason = vbNullString
 1364:     InvalidateTangent heldReason
 1365:     P3ForceTangentRebuild = False
 1366:     Exit Function
 1367:   End If
 1368:   If Not P6CanReuseAssembledFactor() Then
 1369:     InvalidateTangent "FACINV"
 1370:     P3ForceTangentRebuild = False
 1371:     Exit Function
 1372:   End If
 1373:   factorXor = P3FactorPlasticXorCount()
 1374:   If localIteration <= 0 Then
 1375:     If P3ActivePlasticPointCount <> 0 Or factorXor <> 0 Or P3JointContactChangedSinceFactor() Or P3AnyTangentDirty() Then
 1376:       InvalidateTangent "FIRST"
 1377:       P3ForceTangentRebuild = False
 1378:       Exit Function
 1379:     End If
 1380:     P3ShouldRebuildTangent = False
 1381:     Exit Function
 1382:   End If
```

### FEMEngine.bas:1462–1477

```vb
 1462:   If P3PrevResidualNorm > 0# Then
 1463:     If residualNow > 0.7 * P3PrevResidualNorm Then
 1464:       InvalidateTangent "RESID"
 1465:       P3ForceTangentRebuild = False
 1466:       Exit Function
 1467:     End If
 1468:   End If
 1469:   P3TangentStaleIters = P3TangentStaleIters + 1
 1470:   If P3TangentStaleIters >= 3 Then
 1471:     P3TangentStaleIters = 0
 1472:     InvalidateTangent "STAG"
 1473:     P3ForceTangentRebuild = False
 1474:     Exit Function
 1475:   End If
 1476:   P3ShouldRebuildTangent = False
 1477: End Function
```

### FEMEngine.bas:1507–1522

```vb
 1507: Private Sub P3AcceptDampedLineSearch(ByVal alpha As Double, ByVal residualBefore As Double)
 1508:   Dim q As Double
 1509:   Dim growing As Boolean
 1510:   P6LsNoteAccept alpha, True
 1511:   P6LsNoteQ residualBefore, ResidualNormFree
 1512:   q = 2#
 1513:   If residualBefore > 0# Then q = ResidualNormFree / residualBefore
 1514:   growing = (residualBefore > 0# And ResidualNormFree > residualBefore)
 1515:   If CanReuseFactorAfterLineSearch(P3NewtonPolicy, alpha, q, growing, P3AfterCutback, P3ActiveSetForcesRebuild(), P6CanReuseAssembledFactor()) Then
 1516:     P3ForceTangentRebuild = False
 1517:     P3ForceRebuildReason = vbNullString
 1518:     P3Ls50Pending = True
 1519:   Else
 1520:     P3RequestTangentRebuild "LINESEARCH_OK"
 1521:   End If
 1522: End Sub
```

### FEMEngine.bas:2480–2534

```vb
 2480:       P6PhysStep = stepSize
 2481:       prescribedFactor = P3PrescribedFactor(stageId, targetFactor)
 2482:       P3CopyCommittedDispToTrial
 2483:       For i = 0 To lastDof
 2484:         incrementBase(i) = P3CommittedDisp(i)
 2485:       Next i
 2486:       If P6MixedUP Then P6UpdateContinuityG
 2487:       If stageId = 1 Then
 2488:         selfWeightFactor = targetFactor
 2489:         appliedFactor = 0#
 2490:       Else
 2491:         selfWeightFactor = 0#
 2492:         appliedFactor = targetFactor
 2493:       End If
 2494:       P3BuildTargetForce selfWeightFactor, appliedFactor, targetForce
 2495:       P3ClearTransientFailure
 2496:       localIteration = 0: correctionRatio = 0#: converged = False
 2497:       P6LsWatch = False
 2498:       P3PrevResidualNorm = 0#
 2499:       P3TangentStaleIters = 0
 2500:       lineSearchRebuildUsed = False
 2501:       Do
 2502:         CurrentIncrement = P3SuccessfulIncrementCount + 1
 2503:         CurrentIteration = localIteration
 2504:         If P3TrialStateValid Then
 2505:           P3EvalSkipCount = P3EvalSkipCount + 1
 2506:         Else
 2507:           If Not P3EvaluateTrialState(incrementBase, internalForce) Then Exit Do
 2508:         End If
 2509:         P3UpdateResidualMetrics targetForce, internalForce
 2510:         P3UpdateBoundaryDispError prescribedFactor
 2511:         boundaryOk = P3BoundaryDispSatisfied()
 2512:         If RelativeResidualFree <= P3_ENGINEERING_RESIDUAL And boundaryOk Then
 2513:           If localIteration = 0 Then
 2514:             If Not hasPrescribedDisp Or P3MaxBoundaryDispError <= 0.000000000001 Then
 2515:               converged = True
 2516:               If P3TangentAge = 1 Then P6Age1Success = P6Age1Success + 1
 2517:               Exit Do
 2518:             End If
 2519:           ElseIf correctionRatio <= P3_ENGINEERING_CORRECTION Or RelativeResidualFree <= P3_RESIDUAL_TOLERANCE Then
 2520:             converged = True
 2521:             If P3TangentAge = 1 Then P6Age1Success = P6Age1Success + 1
 2522:             Exit Do
 2523:           End If
 2524:         End If
 2525:         If localIteration >= P3_MAX_GLOBAL_ITERATIONS Then
 2526:           SetAnalysisFailure RESULT_NONCONVERGED, "P3全体反復が上限回数に達しました。Ver=" & FEM_BUILD_STAMP & " 相対残差=" & Format$(RelativeResidualFree, "0.000E+00") & " 補正比=" & Format$(correctionRatio, "0.000E+00") & " 拘束変位誤差=" & Format$(P3MaxBoundaryDispError, "0.000E+00") & "。状態は前増分へ戻します。", vbObjectError + 3202, FailureElement, FailureGaussPoint, CurrentIncrement, CurrentIteration
 2527:           Exit Do
 2528:         End If
 2529:         For i = 0 To lastDof
 2530:           Disp(i) = 0#
 2531:           If NodeCond(i) <> 0 Then Disp(i) = P3TargetBoundaryDisp(i, prescribedFactor) - TDisp(i)
 2532:           Force(i) = targetForce(i) - internalForce(i)
 2533:         Next i
 2534:         If P3ShouldRebuildTangent(localIteration, ResidualNormFree) Then
```

### FEMEngine.bas:2598–2610

```vb
 2598:         If localIteration <= 2 Then
 2599:           stepSize = stepSize * 1.5
 2600:         ElseIf localIteration <= 5 Then
 2601:           stepSize = stepSize
 2602:         ElseIf localIteration <= 8 Then
 2603:           stepSize = stepSize * 0.7
 2604:         Else
 2605:           stepSize = stepSize * 0.5
 2606:         End If
 2607:         If stepSize > P3_MAX_STEP_SIZE Then stepSize = P3_MAX_STEP_SIZE
 2608:         If stepSize > 1# - stageFactor And (1# - stageFactor) > 0# Then stepSize = 1# - stageFactor
 2609:         If stepSize < P3_MIN_STEP_FACTOR Then stepSize = P3_MIN_STEP_FACTOR
 2610:         P3LastSuccessStepSize = stepSize
```

### FEMSolver.bas:457–472

```vb
  457:   ' 本番ソルバはBAND固定。CSR/Uzawa/混合はコードに残すがここでは起動しない。
  458:   P6SolverPolicy = "BAND"
  459:   P6UseCSR = False
  460:   P6SolverMode = "BAND_LU"
  461:   P6MixedCoupledKrylov = False
  462: 
  463:   memoryLimitMB = P6ReadSetting("SOLVER_MEMORY_LIMIT_MB", 512#)
  464:   If memoryLimitMB < 1# Then memoryLimitMB = 512#
  465:   P6SolverMemoryLimitBytes = memoryLimitMB * 1024# * 1024#
  466:   P6IterativeTolerance = P6ReadSetting("SOLVER_TOLERANCE", 0.00000001)
  467:   If P6IterativeTolerance <= 0# Then P6IterativeTolerance = 0.00000001
  468:   P6IterativeMaxIterations = CLng(P6ReadSetting("SOLVER_MAX_ITERATIONS", 2000#))
  469:   If P6IterativeMaxIterations < 1 Then P6IterativeMaxIterations = 2000
  470:   requestedRestart = CLng(P6ReadSetting("SOLVER_GMRES_RESTART", 30#))
  471:   P6GMRESRestartRequested = requestedRestart
  472:   P6GMRESRestart = P6ChooseGMRESRestart(requestedRestart)
```

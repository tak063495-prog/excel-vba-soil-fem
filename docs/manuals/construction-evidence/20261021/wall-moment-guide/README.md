# 連続体擁壁の断面力：既存工程3の後処理例

説明書の第5章59–65ページ。鋭い入隅の応力特異性は細分化だけでは解消しないこと、局所最大応力と断面N/V/Mの違い、曲げを取り出すための厚さ方向の分割と固定断面、出力設定・符号・積分法を説明します。

## 元の実計算と適用範囲

- [粗メッシュ wall_davis08](../../20261020/wall_davis08/result.json.gz)：1,857節点・580 Q8、壁40要素。
- [細メッシュ wall_davis12](../../20261020/wall_davis12/result.json.gz)：3,383節点・1,074 Q8、壁68要素。
- 工程3 BIRTH（背面盛土完了）、Fs=1、壁材料番号2。工程4 SRM下端とは別の状態です。
- 壁厚0.4m、奥行1m、E=25,000,000kPa、ν=0.2、γ=24kN/m³。既存の倍精度・収束条件・材料則は変更していません。新規のExcel計算ではなく保存4GP応力の再整理です。
- 保存された圧縮正テンソルCをT=-Cへ変換。縦壁n=+Y、底盤n=+X。引張N正、Mz反時計回り正。N/VはkN/m、MはkN·m/m。

## 再現

リポジトリのルートで、Python・NumPy・Matplotlib（図の日本語フォントはYu Gothic）を使います。既存の`tests/check_manual_case.py`のQ8形状関数を参照します。Excelは起動しません。

```powershell
python docs/manuals/construction-evidence/20261021/wall-moment-guide/rebuild_sections.py --out tmp/wall-moment-review
python docs/manuals/construction-evidence/20261021/wall-moment-guide/rebuild_figures.py --out tmp/wall-moment-review
```

`sections.csv`は19同一座標断面×2メッシュ、両側値・差・12/24点積分差を保存。`summary.json`は元データSHA、4解析解検算（X/Y断面×厚さ1/3分割）と局所釣合い診断を含みます。Q8境界との交点・逆写像で断面を分割し、4GP応力を1,r,s,rsで近似復元しています。断面は±1e-7m側の平均。4GPからの近似復元は厳密な節点内力や接触反力ではありません。

**設計用に収束確認済みの曲げ値ではありません。** 細メッシュの底盤x=18.0のM=-10.566kN·m/mに対し、左側部分の復元応力＋自重のM不釣合いは-18.955kN·m/m。粗密の断面値が近いだけでは採用できません。元ソルバーの節点残差と、この後処理の局所釣合い診断は異なります。診断の4辺は同じ復元応力から求めており、独立に既知の境界反力による証明ではありません。

`actual_meshes.png`は実計算の入力形状、`moment_comparison.png`は保存工程3の後処理値、`mesh_plan_schematic.png`は厚さ2/4/8分割を比較する**計画の模式図**です。最後の図は実行した解析結果でも精度保証でもありません。厚さ方向を整えた複数段階のメッシュ、同じ座標での断面比較、節点内力・境界反力との照合を追加してから部材設計へ使用します。

`manual_section_source.py`は既存PDF組版への追加ページ定義で、単独の全冊ビルダーではありません。本文・表・図は今回のCSV/JSONと対応させています。

## 参考

[S9 Ansys — Assessing Stress Singularity](https://innovationspace.ansys.com/courses/courses/tips-and-tricks-for-structural-simulation-in-ansys-mechanical/lessons/assessing-stress-singularity/)。ピークと応力積分の区別を参照。圧力容器の応力分類法を擁壁の設計規準へ転用するものではありません。

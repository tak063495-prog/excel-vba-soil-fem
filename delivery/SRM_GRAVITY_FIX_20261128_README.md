# 1m版 自重収束修正版

[Excel・DLL・検証記録一式](2DSoilFEM_1m_gravity_fix_c6000_20261128.zip?raw=true)／[修正内容と検証](../docs/SRM_GRAVITY_FIX_20261128.md)

材料11・12をc=6,000kPaとした検証用入力です。その他の入力・設定・最終収束判定は維持し、自重解析が荷重係数1.0まで収束したことを確認しました（8増分、約9.7分）。今回、1m版のSRM安全率は未評価です。

ZIPを展開し、Excelと`soilfem_band_fast_x64.dll`を同じフォルダーに置き、64ビット版Excelで開いてください。外部最適化ソフトは不要です。

ブック SHA256: `0d866b9b2e8fba9f63cb253d9530f0cbb8a6c242e25008643bb1a592651fb221`

DLL SHA256: `5608e5b3c895162fbbb57b5ed4722e17e8c936a28d70679f02526658720d151f`

ZIP SHA256: `52da716f864763c7c1b1a063a9f75a92460e2e0cee087982190e8996868b06ed`

Issue #73・#75の修正を含む更新版です。[修正と実機検証](../docs/ISSUES_73_75_20261128.md)を参照してください。

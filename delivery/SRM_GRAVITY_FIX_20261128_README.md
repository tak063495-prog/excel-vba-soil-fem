# 1m版 自重収束修正版

[Excel・DLL・検証記録一式](2DSoilFEM_1m_gravity_fix_c6000_20261128.zip?raw=true)／[修正内容と検証](../docs/SRM_GRAVITY_FIX_20261128.md)

材料11・12をc=6,000kPaとした検証用入力です。その他の入力・設定・最終収束判定は維持し、自重解析が荷重係数1.0まで収束したことを確認しました（8増分、約9.1分）。今回、1m版のSRM安全率は未評価です。

ZIPを展開し、Excelと`soilfem_band_fast_x64.dll`を同じフォルダーに置き、64ビット版Excelで開いてください。外部最適化ソフトは不要です。

ブック SHA256: `1941e8f6f794f749031ed1a2ecd167a1b965dd76b8f991e48905c07d1a23bcfb`

DLL SHA256: `5608e5b3c895162fbbb57b5ed4722e17e8c936a28d70679f02526658720d151f`

ZIP SHA256: `b61bb0d22c2c51ba002e1a7ad4a3c2cc450f0cd215836fcac2cb0839959b90a1`

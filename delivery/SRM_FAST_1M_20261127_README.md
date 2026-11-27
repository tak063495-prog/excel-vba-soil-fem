# 1m版 SRM高速化・issue修正版

[修正版一式（ZIP）](2DSoilFEM_1m_SRM_fast_issuefix_20261127.zip?raw=true)

前回配布した修正版一式と同じファイルです。ZIPを展開し、Excelと
`soilfem_band_fast_x64.dll` を同じフォルダーに置いて、64ビット版Excelで使用してください。
ソース、使い方、検証記録、issue対応状況も同梱しています。

入力値・解析条件・収束判定は前の1m修正版から維持しています。
同じSRM開始確認地点までの実測は88.2分から5.5分（約16倍）でした。
全SRM探索の完了時間ではなく、最終安全率は計算・保存していません。

issue #62～67と、提供されたphase-Aブックの#69を修正しています。
#70・71は既存PR #68内の修正コードを確認したもので、この配布物にその追加機能は含みません。

関連修正：[PR #72](https://github.com/tak063495-prog/excel-vba-soil-fem/pull/72)。
mainへのマージは行っていません。

ZIP SHA256: `69a537f73957eee8196df4d3d0781994995624bb7a566f9b92f45b2944f8a89b`

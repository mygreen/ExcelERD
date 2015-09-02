# ExcelERD

ExcelERDのExcel2007対応版。

http://typea.info/tools/exerd/manual/


主な変更点は以下の通り。
- Excel2007以上に対応。
-- Excel2007/2010で動作確認しています。
- ER図を出力する際のオプションを追加。
-- 物理名・論理名の同時出力対応。

## Lisence

Apache2.0

## 使い方

https://mygreen.github.com/ExcelERD/index.html

# ソースの管理方法

VBAのライブラリAriawaseを使用し、VBAマクロのソースコード抽出や、既存のファイルのバージョンアップを行う。

## VBAの抽出
1. 「bin」フォルダに、抽出対象のExcelファイルを格納する。
2. バッチファイル「vba_export.bat」を実行する。
3. バッチファイルを実行すると「src」フォルダに抽出される。

## VBAの取り込み
1. 「src」フォルダに、取り込み対象のマクロファイルを格納する。
2. 「bin」フォルダに、取り込み先のExcelファイルを配置する。
2. バッチファイル「vba_import.bat」を実行する。
3. バッチファイルを実行すると「bin/<Excelファイル>」に取り込みされる。





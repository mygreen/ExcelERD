@ECHO OFF

%~d0
CD %~p0

ECHO ExcelERDのVBAマクロを抽出します。

SET FILENAME=ExERD.xlsm
IF NOT EXIST "bin\%FILENAME%" (
    ECHO 抽出元のファイル bin\%FILENAME% が見つかりません。
    GOTO FINISH
)

RMDIR /s /q src\%FILENAME%

cscript vbac.wsf decombine

:FINISH
PAUSE

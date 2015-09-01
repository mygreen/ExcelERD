@ECHO OFF

%~d0
CD %~p0

ECHO ExcelERDのVBAマクロを取り込みします。

SET FILENAME=ExERD.xls
IF NOT EXIST "bin\%FILENAME%" (
    ECHO 取り込み先のファイル bin\%FILENAME% が見つかりません。
    GOTO FINISH

)

cscript //nologo vbac.wsf combine

:FINISH
PAUSE

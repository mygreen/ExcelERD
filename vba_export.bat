@ECHO OFF

%~d0
CD %~p0

ECHO ExcelERD��VBA�}�N���𒊏o���܂��B

SET FILENAME=ExERD.xlsm
IF NOT EXIST "bin\%FILENAME%" (
    ECHO ���o���̃t�@�C�� bin\%FILENAME% ��������܂���B
    GOTO FINISH
)

RMDIR /s /q src\%FILENAME%

cscript vbac.wsf decombine

:FINISH
PAUSE

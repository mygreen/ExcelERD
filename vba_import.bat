@ECHO OFF

%~d0
CD %~p0

ECHO ExcelERD��VBA�}�N������荞�݂��܂��B

SET FILENAME=ExERD.xls
IF NOT EXIST "bin\%FILENAME%" (
    ECHO ��荞�ݐ�̃t�@�C�� bin\%FILENAME% ��������܂���B
    GOTO FINISH

)

cscript //nologo vbac.wsf combine

:FINISH
PAUSE

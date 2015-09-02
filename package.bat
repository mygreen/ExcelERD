@ECHO OFF

%~d0
CD %~p0

ant -f build.xml

PAUSE

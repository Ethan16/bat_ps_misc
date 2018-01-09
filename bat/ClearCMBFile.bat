@echo off

set path=%path%;G:/
FOR /D %%i IN ($SD_*) DO rd /q /s %%i

pause

@echo off
netstat -a -n >a.txt
type a.txt|find "7626" &&echo "Congratuations!You have infected Glacier!"
del a.txt
pause & exit
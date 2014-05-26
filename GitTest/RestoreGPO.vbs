REM **** Restore **********************************************

dim x , a

set x = CreateObject("GPExplorer.PolicyManager.1")

REM ***** RestorePolicy(NULL, Backup folder path, Domain Name, DC Name, option)

x.RestorePolicy NULL, "C:\GPO Backups\{F913DA30-DACA-4BD4-8BD3-C728C89FE927}","gpdom100.lab","larandiat-dv01.gpdom100.lab", 0

set x = nothing

Wscript.Echo "Restore Completed"

REM ***********************************************************
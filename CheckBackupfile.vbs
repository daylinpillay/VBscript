Option Explicit

Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")

'WScript.Echo (Year(Now) & Month(Now) & Day(Now-1) & 1800 & ".BAK")


If (fso.FileExists("\\olympic\Backups\titanic\Pushed\LiveData_db_" & Year(Now) & Month(Now)& Day(Now-1) & 1800 & ".BAK")) Then
     'This is where you send your email and attach log file
     WScript.Echo("Your file exists.")

'If (fso.FileExists("C:\backup\LiveData_db_" & Year(Now) & Month(Now)& Day(Now-1) & 1800 & ".BAK")) Then
'     'This is where you send your email and attach log file
'     WScript.Echo("Your file exists.")
Else
     'Exit script if log file does not exist
     WScript.Echo("Your file is not there")
End If


'Example filename: LiveData_db_201611151800.BAK

'\\olympic\Backups\titanic\Pushed
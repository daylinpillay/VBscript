Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("C:\Program Files (x86)\Teamviewer\TeamViewer10_Logfile.log", ForReading)

For i = 1 to 5
    objTextFile.ReadLine
Next

strLine = objTextFile.ReadLine
Wscript.Echo strLine

objTextFile.Close
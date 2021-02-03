Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    ("C:\Program Files (x86)\Teamviewer\TeamViewer10_Logfile.log", ForReading)

Do Until objTextFile.AtEndOfStream
    strComputer = objTextFile.ReadLine
    Wscript.Echo strComputer
Loop

objTextFile.Close

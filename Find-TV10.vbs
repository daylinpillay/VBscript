' Declare the constants
Dim oFSO
Const HKLM = &H80000002 ' HKEY_LOCAL_MACHINE
'Const REG_SZ = 1 ' String value in registry (Not DWORD)
Const ForReading = 1 
Const ForWriting = 2

' Set File objects...
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objDictionary = CreateObject("Scripting.Dictionary")
Set objDictionary2 = CreateObject("Scripting.Dictionary")

' Set string variables
strDomain = "mistras.local" ' Your Domain
strPCsFile = "DomainPCs.txt" 
strPath = "C:\logs\" ' Create this folder
strWorkstationID = "C:\logs\WorkstationID.txt"

If objFSO.FolderExists(strPath) Then
Wscript.Echo "This program will collect Workstation ID on remote compter(s)"
Else
Wscript.Echo "This program will collect Workstation ID on remote compter(s)"
oFSO.CreateFolder strPath
End If

' Get list of domain PCs - Using above variables.
strMbox = MsgBox("Would you like info for entire domain?",3,"Hostname")

'an answer of yes will return a value of 6, causing script to collect domain PC info
If strMbox = 6 Then
Set objPCTXTFile = objFSO.OpenTextFile(strPath & strPCsFile, ForWriting, True)
Set objDomain = GetObject("WinNT://" & strDomain) ' Note LDAP does not work
objDomain.Filter = Array("Computer")
For Each pcObject In objDomain
objPCTXTFile.WriteLine pcObject.Name
Next
objPCTXTFile.close

Else
'an answer of no will prompt user to input name of computer to scan and create PC file
strHost = InputBox("Enter the computer you wish to get Workstation ID","Hostname"," ")
Set strFile = objfso.CreateTextFile(strPath & strPCsFile, True)
strFile.WriteLine(strHost)
strFile.Close
End If


' Read list of computers from strPCsFile into objDictionary
Set readPCFile = objFSO.OpenTextFile(strPath & strPCsFile, ForReading)
i = 0
Do Until readPCFile.AtEndOfStream 
strNextLine = readPCFile.Readline
objDictionary.Add i, strNextLine
i = i + 1
Loop
readPCFile.Close


' Build up the filename found in the strPath
strFileName = "Workstation ID_" _
& year(date()) & right("0" & month(date()),2) _
& right("0" & day(date()),2) & ".txt"

' Write each PC's software info file...
Set objTextFile2 = objFSO.OpenTextFile(strPath & strFileName, ForWriting, True)

For each DomainPC in objDictionary
strComputer = objDictionary.Item(DomainPC)

Set wmi = GetObject("winmgmts://./root/cimv2")
qry = "SELECT * FROM Win32_PingStatus WHERE Address='" & strComputer & "'"
For Each response In wmi.ExecQuery(qry)
  If IsObject(response) Then
    hostAvailable = (response.StatusCode = 0)
  Else
    hostAvailable = False
  End If
Next


On error resume Next

If hostAvailable Then
  'check for TeamViewer ID

' WMI connection to the operating system note StdRegProv
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
strComputer & "\root\default:StdRegProv")

' These paths are used in the filenames you see in the strPath
pcName = "SYSTEM\CurrentControlSet\Control\ComputerName\ActiveComputerName\"
pcNameValueName = "ComputerName"
objReg.GetStringValue HKLM,pcName,pcNameValueName,pcValue
strKeyPath = "SOFTWARE\Wow6432Node\TeamViewer\Version10\"
strValueName = "ClientID"
objReg.GetDWORDValue HKLM,strKeyPath, strValueName, strValue

If IsNull(strValue) Then
	strKeyPath = "SOFTWARE\Wow6432Node\TeamViewer\Version10\"
	strValueName = "ClientID"
	objReg.GetDWORDValue HKLM,strKeyPath,strValueName,strValue
End If



If IsNull(strValue) Then
	strValue = " No Teamviewer ID"
End If

Set objReg = Nothing
Set ObjFileSystem = Nothing

objTextFile2.WriteLine(vbCRLF & "==============================" & vbCRLF & _
"Current Workstation ID: " & UCASE(strComputer) & vbCRLF & Time & vbCRLF & Date _
& vbCRLF & "Teamviewer ID:" & "" & strValue & vbCRLF _
& "----------------------------------------" & vbCRLF)

'GetWorkstationID()
strValue = NULL

Else

  'remote host unavailable

End If
Next

WScript.echo "Finished Scanning Network check : " & strPath

'objFSO.DeleteFile(strWorkstationID)
objFSO.DeleteFile(strPath & strPCsFile)
wscript.Quit


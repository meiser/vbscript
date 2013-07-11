Const ForReading = 1

'Regex Office 2010 (Office 14)
Set RegOffice2010 = New RegExp
RegOffice2010.IgnoreCase = True
RegOffice2010.Global = True
RegOffice2010.Pattern = "\S*office14\S*"

'Regex Office 2007 (Office12)
Set RegOffice2007 = New RegExp
RegOffice2007.IgnoreCase = True
RegOffice2007.Global = True
RegOffice2007.Pattern = "\S*office12\S*"

'Regex Office 2003 (Office 11)
Set RegOffice2003 = New RegExp
RegOffice2003.IgnoreCase = True
RegOffice2003.Global = True
RegOffice2003.Pattern = "\S*office11\S*"


'get current directory
Dim shl
Set shl = WScript.CreateObject("WScript.Shell")
tempDir = shl.ExpandEnvironmentStrings("%SystemRoot%") & "\Temp\"



'objStartFolder = shl.CurrentDirectory

ProduKeyFile = """\\oelit12\C$\Program Files (x86)\bsAG\BMS\Client\ProduKey.exe"""

'ProduKeyFile = "\\oelnas01\Work\S.Keller\ProduKey.exe"

'Set wshNetwork = WScript.CreateObject( "WScript.Network" )
'strComputerName = wshNetwork.ComputerName
'DestinationPath = Replace(shl.CurrentDirectory, "C:\", "\\" & strComputerName & "\C$\", 1,1)

Set TypeLib = CreateObject("Scriptlet.TypeLib")
DestinationPath = tempDir & "key_" & TypeLib.Guid & ".txt"

cmd = "cmd /c copy " & ProduKeyFile & " " & tempDir
shl.run cmd,1,TRUE

cmd = "cmd /c " & tempDir & "ProduKey.exe /OfficeKeys 1 /WindowsKeys 0 /IEKeys 0 /SQLKeys 0 /ExchangeKeys 0 /scomma """" >> " & DestinationPath
shl.run cmd,1,TRUE

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(DestinationPath,ForReading)

office2010 = ""
office2007 = ""
office2003 = ""
do while not objFile.AtEndOfStream 
    strLine =  objFile.ReadLine()
    program = Split(strLine, ",")
	Set officeMatches = RegOffice2010.Execute(program(3))

	if officeMatches.Count > 0 then
		'Wscript.Echo "Office 2010", program(3)
		
		if office2010 = "" then
			office2010 = program(0)+" "+program(2)
		else
			office2010 = office2010 & ", " & program(0)+" "+program(2)
		end if
	end if
	
	Set officeMatches = RegOffice2007.Execute(program(3))

	if officeMatches.Count > 0 then
		Wscript.Echo "Office 2007", program(3)
		if office2007 = "" then
			office2007 = program(0)+" "+program(2)
		else
			office2007 = office2007 & ", " & program(0)+" "+program(2)
		end if
	end if
	
	Set officeMatches = RegOffice2003.Execute(program(3))

	if officeMatches.Count > 0 then
		Wscript.Echo "Office 2003", program(3)
		if office2003 = "" then
			office2003 = program(0)+" "+program(2)
		else
			office2003 = office2003 & ", " & program(0)+" "+program(2)
		end if
	end if
    
loop

Set objTextFile = objFSO.CreateTextFile("office2003.txt", True)
objTextFile.Write(office2003)
objTextFile.Close

Set objTextFile = objFSO.CreateTextFile("office2007.txt", True)
objTextFile.Write(office2007)
objTextFile.Close

Set objTextFile = objFSO.CreateTextFile("office2010.txt", True)
objTextFile.Write(office2010)
objTextFile.Close

Office2013KeyScript = "C:\Program Files (x86)\Microsoft Office\Office14\OSPP.VBS"
Office2013DestinationPath = tempDir & "office2013_" & TypeLib.Guid & ".txt"

'Wscript.Echo Office2013DestinationPath

'If objFSO.FileExists(Office2013KeyScript) Then
'	cmd = "cmd /c cscript """ & Office2013KeyScript & """ /dstatus >> " & Office2013DestinationPath
	
'	Wscript.Echo cmd 
'	shl.run cmd,1,TRUE

'	Set objFile = objFSO.OpenTextFile(Office2013DestinationPath,ForReading)
'	do while not objFile.AtEndOfStream
'		strLine =  objFile.ReadLine()
'		program = Split(strLine, ":")
'		Wscript.Echo strLine
		'if i > 5 and i < 12 then
		'Last 5 characters of installed product key:	
			'OutputOffice2013 = OutputOffice2013 & "<Property name=""" & program(0) & """ value=""" & program(1) & """/>"
		'end if
'	loop

'end if


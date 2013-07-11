'Dim objNetwork
'Set objNetwork = CreateObject("WScript.Network")
'strUserName = objNetwork.UserName
'MsgBox strUserName

Const ForReading = 1

'Regex Office
Set myRegExp = New RegExp
myRegExp.IgnoreCase = True
myRegExp.Global = True
myRegExp.Pattern = "\S*office\S*"

'get current directory
Dim shl
Set shl = WScript.CreateObject("WScript.Shell")
'objStartFolder = shl.CurrentDirectory

'ProduKeyFile = """\\oelit12\C$\Program Files (x86)\bsAG\BMS\Client\ProduKey.exe"""

'ProduKeyFile = "\\oelnas01\Public\Speiseplan\ProduKey"
ProduKeyFile = "\\oelit12\c$\Program Files (x86)\bsAG\BMS\Client\ProduKey.exe"

Set objFSO = CreateObject("Scripting.FileSystemObject")

'Set wshNetwork = WScript.CreateObject( "WScript.Network" )
'strComputerName = wshNetwork.ComputerName
'DestinationPath = Replace(shl.CurrentDirectory, "C:\", "\\" & strComputerName & "\C$\", 1,1)

tempDir = shl.ExpandEnvironmentStrings("%SystemRoot%") & "\Temp\"

Set TypeLib = CreateObject("Scriptlet.TypeLib")
DestinationPath = tempDir  & TypeLib.Guid & ".txt"


If objFSO.FileExists(ProduKeyFile) Then
	objFSO.CopyFile ProduKeyFile, tempDir 

	cmd = "cmd /c " & tempDir & "ProduKey.exe /scomma """" >> " & DestinationPath
	shl.run cmd,0,TRUE
else
	Wscript.Echo "nix da"
end if

'fso.CopyFile "\\oelnas01\Daten\EDV\", shl.CurrentDirectory

Set objFile = objFSO.OpenTextFile(DestinationPath,ForReading)
Output = "<SubNode name=""Microsoft Office Pakete"">"
OutputOther = "<SubNode name=""andere Microsoft Produkte"">"
do while not objFile.AtEndOfStream
	strLine =  objFile.ReadLine()
    program = Split(strLine, ",")
	Set officeMatches = myRegExp.Execute(program(0))
	if officeMatches.Count > 0 then
		Output = Output & "<Property name=""" & program(0) & """ value=""" & program(2) & """/>"
	else
		OutputOther = OutputOther & "<Property name=""" & program(0) & """ value=""" & program(2) & """/>"
	end if
loop

'objFile.Close

Output = Output & "</SubNode>"
OutputOther = OutputOther & "</SubNode>"

Set OutputFile = CreateObject ("Scripting.FileSystemObject")
Set OutputStream = OutputFile.CreateTextFile ("lizenzierungresult.xml", True)


OutputStream.WriteLine Output+OutputOther
OutputStream.Close

'objFSO.DeleteFile DestinationPath
'objFSO.DeleteFile "C:\Windows\Temp\ProduKey.exe"

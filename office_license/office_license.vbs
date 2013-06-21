Const ForReading = 1

'Regex Office
Set myRegExp = New RegExp
myRegExp.IgnoreCase = True
myRegExp.Global = True
myRegExp.Pattern = "\S*office\S*"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("key.txt",ForReading)
do while not objFile.AtEndOfStream 
    strLine =  objFile.ReadLine()
    program = Split(strLine, ",")
	Set officeMatches = myRegExp.Execute(program(0))
	
	if officeMatches.Count > 0 then
		Set objTextFile = objFSO.CreateTextFile(objFile.Line-1, True)
		objTextFile.Write(program(0)+" "+program(2))
	end if
    objTextFile.Close
loop
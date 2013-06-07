Dim account_id 
account_id  = InputBox("Bitte Import Empfänger Email Angeben (z.B. Mustermann, Max:f123456789ef2332322322)")

Dim shl
Set shl = WScript.CreateObject("WScript.Shell")

Set objFSO = CreateObject("Scripting.FileSystemObject")
objStartFolder = shl.CurrentDirectory+"\pst\"
objReadyFolder = shl.CurrentDirectory+"\pst\importiert\"

If Not objFSO.FolderExists(objStartFolder) Then
	objFSO.CreateFolder(objStartFolder)
End If

If Not objFSO.FolderExists(objReadyFolder) Then
    	objFSO.CreateFolder(objReadyFolder)
End If

Set objFolder = objFSO.GetFolder(objStartFolder)

Set colFiles = objFolder.Files

For Each objFile in colFiles
    'Wscript.Echo objFile.Path
	If UCase(objFSO.GetExtensionName(objFile.name)) = "PST" Then
    	cmd = """" &"C:\Program Files (x86)\EASY xBASE\xadmin\xbarchpst"&"""" &"-vvv -f -l "&"""" &"PstImportLog\"& objFile.Name&".txt" &"""" &" -r auto " &""""&"oelex11" &""""& " " & """"& account_id &"""" & " " &""""& "xBASE" &"""" & " " &""""& objFile.Path&""""
		'Wscript.Echo cmd
		shl.Run cmd,1, TRUE
		objFSO.MoveFile objFile.Path, objReadyFolder
    End If
Next
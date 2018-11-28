<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/objects/cFileUploader.asp"-->
<!-- #include virtual="/objects/cUploadedFile.asp"-->
<%
Dim upFile, FileSys, FilePath, theFileName
Dim strUploadPath, strProofFolder
strUploadPath = "c:\Websites\domain\secure\proofs\<photographer>\" & strProofFolder
Dim oUploader
Set oUploader = New cFileUploader
oUploader.Upload()


Private Function AddFileAttributesToDb(filename,filepath,id)
	Dim cmd
	Set cmd = Server.CreateObject("ADODB.Command")	
	With cmd
		.ActiveConnection = CONNECTION_STRING
		.CommandText = "sp__AddProductFile"
		.CommandType = adCmdStoredProc	
		.Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)				
		.Parameters.Append .CreateParameter("@fileName",adVarChar,adParamInput,200,filename)
		.Parameters.Append .CreateParameter("@filePath",adVarChar,adParamInput,255,filepath)		
		.Parameters.Append .CreateParameter("@fileId",adInteger, adParamOutput, ,id)						
		.Execute , ,adExecuteNoRecords
		id = .Parameters("@fileId")
		'm_ReturnCode = .Parameters("Return")
		AddFileAttributesToDb = CBool(.Parameters("Return"))
	End With
	CloseCmd(cmd)	
End Function
	
'Set oFS = Server.CreateObject("Scripting.FileSystemObject")
'echobr(oFS.FolderExists(strUploadPath))


Dim urlArgs, errorCode
	
' Check if any files were uploaded
If oUploader.Files.Count > 0 Then
	
	' Loop through the uploaded files
	For Each upFile In oUploader.Files.Items	
		FilePath = strUploadPath & "\" & upFile.FileName
		Set FileSys = CreateObject("Scripting.FileSystemObject")
		upFile.SaveToDisk strUploadPath

		Dim id, errorCount
		errorCount = 0
		id = 0
		blnSuccess = AddFileAttributesToDb(upFile.FileName,strUploadPath,id)

		If blnSuccess = False Then
			errorCode = 1000
			errorCount = errorCount + 1
			Set oFS = Server.CreateObject("Scripting.FileSystemObject")
			oFS.DeleteFile(FilePath)
		End If
	Next
	
	Set oUploader = Nothing
	If errorCount > 0 Then
		urlArgs = "?err=" & errorCount & "&code=" & errorCode
	Else
		urlArgs = ""
	End If
	PageRedirect("/cms/prodfile.asp" & urlArgs)
End If


'Dim oFile 
'Set oFile = New cUploadedFile
'oFile.ID = 0
'oFile.FileName = fname
'oFile.FilePath = fpath
'oFile.SaveFileToDisk()
'oFile.SaveToDisk()
'Set oFile = Nothing

%>

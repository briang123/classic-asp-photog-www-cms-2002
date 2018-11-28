<MM:BeginLock translatorClass="MM_ASPSCRIPT" type="script" depFiles="" orig="%3C%25
'********************************************************************************
' START IMAGE CONFIGURATION SECTION
'********************************************************************************
'
' MODULE: 		__fsoConfig.asp
' CREATED BY: 	Brian Gaines
' CREATED ON: 	02/27/2003
'
' PURPOSE: 			The following module formats the Image Management Console
'					These values are currently hardcoded, but can be retrieved
'					from a database as a future enhancement. We would need to 
'					set up a config window for the Admin user to modify/save.

' User-defined variables that need to be set by the user.
Dim 		strRootDirName, _
			strAdmTopDir, _
			strImgTopDir, _
			strFileTopDir, _
			strRootWebServer, _
			strFileExtHeader, _
			blnCanDelete, _
			strWebImgDir, _
			strRootPath, _
			strBaseWebPath, _
			strCeilingFolder, _
			strHeaderImage, _
			strComboTopLevel, _
			strFileCeilingFolder, _
			strFileComboTopLevel, _
			strFileUploadPath, _
			strImgUploadPath, _
			strImageExt
			
Dim objFSO, objFolder

strRootDirName				= %22%22							' The application directory name
strAdmTopDir				= %22\admin\%22						' The top level directory to where the admin pages are stored (LIMITED ACCESS)
strImgTopDir				= %22\Lambent\cms\static\%22	' The file directory after the root application directory where images are stored.
'strImgTopDir				= %22\static\images\%22				' The file directory after the root application directory where images are stored.
strFileTopDir				= %22\Lambent\cms\static\%22	' The file directory after the root application directory where static files (.doc,.pdf,etc.) are stored
strRootWebServer 			= %22C:\Websites\Lambent%22			' The file directory where the application is installed		
strFileExtHeader 			= %22jpg,gif,bmp,doc,zip,pdf,txt,xls,ppt,lam,xml,html,htm%22				' These are the file extensions we want to appear in our file window.
strImageExt					= %22jpg,gif,bmp,png,art,tif%22
blnCanDelete				= True							' Flag that indicates if deletes can be performed on the system.

strWebImgDir = replace(strImgTopDir,%22\%22,%22/%22)				' The web server directory where the images are stored.
If strRootDirName %3C%3E %22%22 Then
	strRootPath = strRootWebServer & %22\%22 & strRootDirName	' The directory path where the application resides
Else
	strRootPath = strRootWebServer
End If

strBaseWebPath 	= %22/%22 & strRootDirName																		' This is the application root web server directory
if strBaseWebPath = %22/%22 Then strBaseWebPath = %22%22

strCeilingFolder = strRootPath & strImgTopDir															' This is the top-most directory that the user can see from the console
If strRootDirName %3C%3E %22%22 Then
	strHeaderImage = %22/%22 & strRootDirName & strWebImgDir & %22images/website/tri.gif%22 	' Header image to be displayed in header
Else
	strHeaderImage = strRootDirName & strWebImgDir & %22images/website/tri.gif%22 	' Header image to be displayed in header
End If
strComboTopLevel = Left(strCeilingFolder,Len(strCeilingFolder)-1)					' The top level directory to match against for the Folders drop-down box. We strip %22\%22
strFileCeilingFolder = strRootPath & strFileTopDir															' The top-most directory that the user can see from the console
strFileComboTopLevel = Left(strFileCeilingFolder,Len(strFileCeilingFolder)-1)	' The top level directory to match against for the Folders drop-down box. We strip %22\%22
strFileUploadPath = replace(strFileComboTopLevel,%22\%22,%22\\%22)									' The path to upload files
strImgUploadPath = replace(strComboTopLevel,%22\%22,%22\\%22)

Dim strWorkflowPath
strWorkflowPath = strImgUploadPath & %22\\Workflow%22

' We will go ahead and open our file system and folder objects
Set objFSO = Server.CreateObject(%22Scripting.FileSystemObject%22)
Set objFolder = objFSO.GetFolder(strCeilingFolder)


If 1 = 2 then
response.write %22strWebImgDir = %22 & strWebImgDir & %22%3Cbr%3E%22
response.write %22strRootPath = %22 & strRootPath & %22%3Cbr%3E%22
response.write %22strBaseWebPath = %22 & strBaseWebPath & %22%3Cbr%3E%22
response.write %22strCeilingFolder = %22 & strCeilingFolder & %22%3Cbr%3E%22
response.write %22strHeaderImage = %22 & strHeaderImage & %22%3Cbr%3E%22
response.write %22strComboTopLevel = %22 & strComboTopLevel & %22%3Cbr%3E%22
response.write %22strFileCeilingFolder = %22 & strFileCeilingFolder & %22%3Cbr%3E%22
response.write %22strFileComboTopLevel = %22 & strFileComboTopLevel & %22%3Cbr%3E%22
response.write %22strFileUploadPath = %22 & strFileUploadPath & %22%3Cbr%3E%22
response.write %22strImgUploadPath = %22 & strImgUploadPath & %22%3Cbr%3E%22
End If
%25%3E" ><MM_ASPSCRIPT><MM:EndLock>
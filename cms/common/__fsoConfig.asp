<%
'********************************************************************************
' START IMAGE CONFIGURATION SECTION
'********************************************************************************
'
' MODULE: 		__fsoConfig.asp
' CREATED BY: 	Brian Gaines
' CREATED ON: 	02/27/2003
'
' PURPOSE: 		The following module formats the Image Management Console
'				These values are currently hardcoded, but can be retrieved
'				from a database. We would need to set up a config window
'				for the Admin user to modify/save.

' User-defined variables that need to be set by the user.
Dim 	strRootDirName, _
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
			strImgUploadPath
			
Dim 	objFSO, objFolder

strRootDirName			= ""				' The application directory name
strAdmTopDir			= "\secure\proofs\julie\"			' The top level directory to where the admin pages are stored (LIMITED ACCESS)
strImgTopDir			= "\secure\proofs\julie\"			' The file directory after the root application directory where images are stored.
strFileTopDir			= "\secure\proofs\julie\"			' The file directory after the root application directory where static files (.doc,.pdf,etc.) are stored
strRootWebServer 		= "C:\Websites\juliestark"						' The file directory where the application is installed		
strFileExtHeader 		= "jpg,gif"							' These are the file extensions we want to appear in our file window.
blnCanDelete			= True								' Flag that indicates if deletes can be performed on the system.

'strRootDirName = ""
'strRootWebServer = "d:\html\users\JulieStarkPhotography\html"

strWebImgDir 			= replace(strImgTopDir,"\","/")								' The web server directory where the images are stored.
strRootPath 			= strRootWebServer & "\" & strRootDirName					' The directory path where the application resides
strBaseWebPath 			= "/" & strRootDirName										' This is the application root web server directory
strCeilingFolder 		= strRootPath & strImgTopDir								' This is the top-most directory that the user can see from the console
strHeaderImage 			= "/" & strRootDirName & "/cms/images/tri.gif" 				' Header image to be displayed in header
strComboTopLevel		= Left(strCeilingFolder,Len(strCeilingFolder)-1)			' The top level directory to match against for the Folders drop-down box. We strip "\"
strFileCeilingFolder	= strRootPath & strFileTopDir								' The top-most directory that the user can see from the console
strFileComboTopLevel	= Left(strFileCeilingFolder,Len(strFileCeilingFolder)-1)	' The top level directory to match against for the Folders drop-down box. We strip "\"
strFileUploadPath		= replace(strFileComboTopLevel,"\","\\")					' The path to upload files
strImgUploadPath		= replace(strComboTopLevel,"\","\\")						' The path to upload images

' We will go ahead and open our file system and folder objects
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strCeilingFolder)
%>
<!--#include virtual="/lambent/cms/lib/__dbfunct.asp"-->
<%
'********************************************************************************
' START COMMON FILE/FOLDER FUNCTIONS
'********************************************************************************
'
' MODULE: __fsoCommon.asp
' CREATED BY: Brian Gaines
' CREATED ON: 02/27/2003
'
' COPYRIGHT: 	Copyright *c*, Brian Gaines, Gaines Consulting
'				All functions created by Gaines Consulting may not be re-used by 
'				other ASP applications without the prior written consent from 
'				Gaines Consulting
'
'********************************************************************************
' START IMAGE DIMENSION FUNCTIONS
'********************************************************************************
'
' COPYRIGHT:	Copyright *c* MM,  Mike Shaffer    
' 				ALL RIGHTS RESERVED WORLDWIDE      
' 				Permission is granted to use this code in your projects, as 
'				long as this copyright notice is included  
'				
' PURPOSE: 		This routine will attempt to identify any filespec passed 
'				as a graphic file (regardless of the extension). This will			
'				work with BMP, GIF, JPG and PNG files. This function gets 
'				a specified number of bytes from any file, starting at the 
'				offset (base 1)
' INPUTS:		flnm - Filespec of file to read
'				offset - Offset at which to start reading
'				bytes - How many bytes to read
function GetBytes(flnm, offset, bytes)

Dim objFSO, objFTemp, objTextStream, lngSize, strBuff, fsoForReading

on error resume next

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	' First, we get the filesize
	Set objFTemp = objFSO.GetFile(flnm)
	lngSize = objFTemp.Size
	Set objFTemp = Nothing
	
	fsoForReading = 1
	Set objTextStream = objFSO.OpenTextFile(flnm, fsoForReading)
	
	if offset > 0 then
		strBuff = objTextStream.Read(offset - 1)
	end if
	
	' Get All
	if bytes = -1 then	
		'ReadAll	
		GetBytes = objTextStream.Read(lngSize)  
	else
		GetBytes = objTextStream.Read(bytes)
	end if
	
	objTextStream.Close
	Set objTextStream = Nothing
	Set objFSO = Nothing

end function

' PURPOSE: 	Functions to convert two bytes to a numeric value (long)
function lngConvert(strTemp)
	lngConvert = clng(asc(left(strTemp, 1)) + ((asc(right(strTemp, 1)) * 256)))
end function

function lngConvert2(strTemp)
 	lngConvert2 = clng(asc(right(strTemp, 1)) + ((asc(left(strTemp, 1)) * 256)))
end function

' PURPOSE: 	This function does most of the real work. It will attempt
'						to read any file, regardless of the extension, and will
'						identify if it is a graphical image.
' INPUTS: 	flnm - Filespec of file to read
'						width - width of image   
'						height - height of image  
'						depth - color depth (in number of colors)
'						strImageType - type of image (e.g. GIF, BMP, etc.)
function gfxSpex(flnm, width, height, depth, strImageType)

	dim strPNG, strGIF, strBMP, strType, strBuff, lngSize, flgFound, strTarget
	dim ExitLoop, lngPos, lngMarkerSize

	strType = ""
	strImageType = "(unknown)"
	
	gfxSpex = False
	
	strPNG = chr(137) & chr(80) & chr(78)
	strGIF = "GIF"
	strBMP = chr(66) & chr(77)
	
	strType = GetBytes(ucase(flnm), 0, 3)

	if strType = strGIF then
	
		strImageType = "GIF"
		Width = lngConvert(GetBytes(flnm, 7, 2))
		Height = lngConvert(GetBytes(flnm, 9, 2))
		Depth = 2 ^ ((asc(GetBytes(flnm, 11, 1)) and 7) + 1)
		gfxSpex = True
	
	elseif left(strType, 2) = strBMP then
	
		strImageType = "BMP"
		Width = lngConvert(GetBytes(flnm, 19, 2))
		Height = lngConvert(GetBytes(flnm, 23, 2))
		Depth = 2 ^ (asc(GetBytes(flnm, 29, 1)))
		gfxSpex = True
	
	elseif strType = strPNG then
	
		strImageType = "PNG"
		Width = lngConvert2(GetBytes(flnm, 19, 2))
		Height = lngConvert2(GetBytes(flnm, 23, 2))
		Depth = getBytes(flnm, 25, 2)
		
		select case asc(right(Depth,1))
		   case 0
			  Depth = 2 ^ (asc(left(Depth, 1)))
			  gfxSpex = True
		   case 2
			  Depth = 2 ^ (asc(left(Depth, 1)) * 3)
			  gfxSpex = True
		   case 3
			  Depth = 2 ^ (asc(left(Depth, 1)))  '8
			  gfxSpex = True
		   case 4
			  Depth = 2 ^ (asc(left(Depth, 1)) * 2)
			  gfxSpex = True
		   case 6
			  Depth = 2 ^ (asc(left(Depth, 1)) * 4)
			  gfxSpex = True
		   case else
			  Depth = -1
		end select
		
	else
		
		' Get all bytes from file
		strBuff = GetBytes(flnm, 0, -1)
		lngSize = len(strBuff)
		flgFound = 0
		
		strTarget = chr(255) & chr(216) & chr(255)
		flgFound = instr(strBuff, strTarget)
		
		if flgFound = 0 then
		   exit function
		end if
		
		strImageType = "JPG"
		lngPos = flgFound + 2
		ExitLoop = false

		do while ExitLoop = False and lngPos < lngSize
		
		   do while asc(mid(strBuff, lngPos, 1)) = 255 and lngPos < lngSize
			  lngPos = lngPos + 1
		   loop
		
		   if asc(mid(strBuff, lngPos, 1)) < 192 or asc(mid(strBuff, lngPos, 1)) > 195 then
			  lngMarkerSize = lngConvert2(mid(strBuff, lngPos + 1, 2))
			  lngPos = lngPos + lngMarkerSize  + 1
		   else
			  ExitLoop = True
		   end if
		
		loop
		
		if ExitLoop = False then
			Width = -1
			Height = -1
			Depth = -1	
		else
			Height = lngConvert2(mid(strBuff, lngPos + 4, 2))
			Width = lngConvert2(mid(strBuff, lngPos + 6, 2))
			Depth = 2 ^ (asc(mid(strBuff, lngPos + 8, 1)) * 8)
			gfxSpex = True	
		end if
	end if

end function


'********************************************************************************
' START FILE SYSTEM OBJECT SUPPORT FUNCTIONS
'********************************************************************************

' PURPOSE:	Resets the directory combo box above the file listing
Sub ReloadFSOComboFolder
	Set objFolder = objFSO.GetFolder(strComboTopLevel)
End Sub	

' PURPOSE: 	This function formats the file size into a value that
'						is a bit more readable. The number is read in as bytes.	
' INPUTS:		number - This is the file size
Function SizeFormat(filesize)
	If filesize < 1000 Then
		SizeFormat = filesize & " b"
	ElseIf filesize > 999 And filesize < 1000000 Then
		SizeFormat = Round(filesize/1000) & " Kb"
	ElseIf filesize > 1000000 Then
		SizeFormat = Round(filesize/1000000) & " Mb"
	End If
End Function

' PURPOSE: 	Converts a path formatted as "C:\\Inetpub\\wwwroot\\..."
'						into a web server path such as "/appfolder/dir1/file1.gif"
' INPUTS:		strPath - Path to where the file is stored.	
Function ConvertToWebPath(strPath)

	Dim normalized, stripRoot
	' reformat the javascript escaping slashes
	normalized = replace(strPath,"\\","\")
	
	' now that we have the format of C:\Inetpub\wwwroot\...
	' we can hide the non-web browser path. where:
	' strRootPath is a user-defined variable that defines the application directory
	' strBaseWebPath is a user-defined variable that represents the base web application path
	stripRoot = replace(normalized,strRootPath,strBaseWebPath) 
	
	' One last formatting technique before we return this value to the caller.
	' change all the backward slashes to forward slashes to be referenced as web server path.
	ConvertToWebPath = replace(stripRoot,"\","/")
	
End Function

' PURPOSE: 	Only display a maximum number of characters of the filename
' INPUTS:		thefile - The file name to check the length of
'						length 	-	The maximum length the file can be.
Function SetMaxFileDisplaySize(thefile,length)
	If Len(thefile) > (length -4) Then 
		SetMaxFileDisplaySize = Left(thefile,length-4) & "<font color=red>???</font>." & 	GetFSOFileExtension(thefile)
	Else
		SetMaxFileDisplaySize = thefile
	End If
End Function

Function CharCount(strSource, strChar)
    Dim intPos, intCount
    intCount = 0
    intPos = InStr(strSource, strChar)
    While intPos
        intCount = intCount + 1
        intPos = InStr(intPos + 1, strSource, strChar)
    Wend
    CharCount = intCount
End Function

' PURPOSE: 	Converts a web server path formatted as "/static/images/"
'						to "C:\Inetpub\wwwroot\static\images\file1.gif"
' INPUTS: 	strPath - Path to convert
Function ConvertToFilePath(strPath)
	ConvertToFilePath = strRootWebServer & replace(strPath,"/","\")
End Function

' PURPOSE:	Get the file extension of the file.
' INPUTS:		strFile - The file
Function GetFSOFileExtension(strFile)
	Dim arrFile, ext	
	If Instr(strFile,".") > 0 Then
		arrFile = split(strFile,".")
		GetFSOFileExtension = arrFile(1)
	Else
		GetFSOFileExtension = "~~~"
	End If
End Function

' PURPOSE:	Validates the operations when the user is attempting to move directories
' INPUTS: 	f1 - Source directory
'						f2 - Target directory	
Function FolderCheck(f1,f2)

	' If the directories do not match, then proceed
	If f1 <> f2 Then 
		FolderCheck = True 
	Else 
		FolderCheck = False
		Response.write "<script>alert('You are not able to perform this operation because the source and \n" & _
										"target directory are pointing to the same directory.');</script>"
	End If
End Function

' PURPOSE:	Checks if a folder exists in the file system
' INPUTS:		directory to check
Function FolderExists(strFolder)
	If objFSO.FolderExists(strFolder) Then
		FolderExists = True
	Else
		FolderExists = False
	End If
End Function

'********************************************************************************
' START FILE SYSTEM OBJECT FUNCTIONS - DIRECTORIES AND FILES
'********************************************************************************

' PURPOSE: 	Builds the combo box with the directory listing. The
'						directory will be selected upone form submission and page
'						reload.
' INPUTS:		fsoFolder - This is the folder object that gets recursed
'						valToMatch - This is the value that we check against when the 
'						form is submitted so we can re-select the folder on page load.
Function BuildFSOFolderCombo(fsoFolder,valToMatch)

	if valToMatch = "" then
		valToMatch = strComboTopLevel
	end if
	
	Dim strDisplayPath

	If fsoFolder = strComboTopLevel Then
		strDisplayPath = "\"
	Else
		strDisplayPath = Replace(fsoFolder.ParentFolder,strComboTopLevel,"") & "\" & fsoFolder.Name
	End If
	
	
	Response.Write("<OPTION VALUE=" & DblQt(fsoFolder) & setComboOption(fsoFolder,valToMatch) & ">" & strDisplayPath & "</OPTION>" & vbCRLF)
					
	' Loop through each sub-folder of the current folder object
	Dim subFolder
	For Each subFolder in fsoFolder.SubFolders
	  	Call BuildFSOFolderCombo(subFolder,valToMatch)
	Next	

End Function


' PURPOSE: 	Get a list of all the files within a sub-directory
'						based on the filtered criteria.
' INPUTS:		aFolder - this is the folder object to get property values for.
Function GetFSOFilesReadOnly(aFolder)

	'----------------------------------------------
	' Database interaction variable declarations
	'----------------------------------------------

	Dim strMode
	Dim strSql
	Dim strFileName
	Dim strFilePath
	Dim blnIsManaged
	Dim intFileId
	Dim intPageId
	Dim strLastModifiedBy
	Dim strIsActive
	Dim strIsLive

	'---------------------------------------------
	' File system variable declarations
	'---------------------------------------------

	Dim fil
	Dim ext
	Dim arrFile
	Dim js_EscapePath
	Dim blnIsFiles
	
	For Each fil In aFolder.files
					

		'-------------------------------------
		' Open recordset object
		'-------------------------------------	
		Call OpenRs(rs)

		' If the current file extension is in the filter list, then we include it.
		If Instr(LCase(strFileExtHeader),LCase(GetFSOFileExtension(fil.Name))) > 0 Then

			strFileName = fil.Name
			strFilePath = ConvertToWebPath(aFolder.Path)
			
			' query database and look for a managed file
			strSql = "SELECT FileId, FileName, FilePath, PageId, FirstName, LastName, f.ActiveFlag, LiveState " & _
					 "FROM tblFiles f, tblUsers u " & _
					 "WHERE f.LastModifiedBy = u.UserId " & _
					 "AND EncryptedFileName = '" & fil.Name & "' " & _
					 "AND FilePath = '" & strFilePath & "'"
					 
			rs.Open strSql, GetConnection
			
			' if no records exist, then the file is not managed
			If rs.BOF And rs.EOF Then
				blnIsManaged = False
			Else
				' get values and store in variables
				blnIsManaged = True
				intFileId = rs(0)
				strFileName = rs(1)
				strFilePath = rs(2)
				intPageId = rs(3)
				strLastModifiedBy = rs(4) & " " & rs(5)
				
				If rs(6) Then
					strIsActive = "Active"
				Else
					strIsActive = "Inactive"
				End If
				
				If rs(7) Then
					strIsLive = "LIVE"
				Else
					strIsLive = "STAGE"
				End If
				
			End If
			
			' we set a flag to indicate our message handler that we have records found
			blnIsFiles = True
			
			' if the file is not managed then we want to mark it
			If Not blnIsManaged Then
			
				strMode = "NEW"
				intFileId = 0
				Response.write("<tr onClick='selectRow(this);' onMouseover=this.bgColor='#8DBCEB' onMouseout=this.bgColor='#FFFFFF' style='cursor:hand'>" & vbCRLf)			

			Else

				strMode = "EDIT"
				' we perform some dhtml features on each row
				Response.write("<tr onClick='selectRow(this);' onMouseover=this.bgColor='#8DBCEB' onMouseout=this.bgColor='#FFFFFF' style='cursor:hand'>" & vbCRLf)

			End If

			Response.write("<td valign=top width=250>")
			If Not blnIsManaged Then
				Response.write("<img src='/lambent/cms/static/uma.gif' border='0' align='absmiddle' alt='Unmanaged file'>&nbsp;")
			Else
				Response.write("<img src='/static/icons/" & GetFSOFileExtension(strFileName) & ".gif' border='0' align='absmiddle' alt='" & strFileName & "'>&nbsp;")
			End If			
			Response.write(strFileName & "</td>" & vbCrLf)
			Response.write("<td valign=middle width=145>" & FormatDate(fil.DateLastModified,"%m/%d/%y-%h:%N %P") & "</td>" & vbCrLf)
			Response.write("<td valign=middle width=130>" & strLastModifiedBy & "</td>" & vbCrLf)
			Response.write("<td valign=middle width=50>" & SizeFormat(fil.size) & "</td>" & vbCrLf)

			' perform an escape of the backslash to support the path to be passed to javascript functions
			js_EscapePath = replace(fil.Path,"\","\\")

			Dim strImgExt
			If intType = 1 Then

				strImgExt = "jpg,png,bmp,gif"
				If Instr(strImgExt,LCase(GetFSOFileExtension(strFileName))) < 1 Then
					Response.write("<td valign=middle width=25 align='center' style=" & DblQt("cursor:hand;") & " onClick=" & DblQt("viewDetails(" & intFileId & ");>&nbsp;&nbsp;<u><font color='#0000FF'>details</font></u></td>" & vbCrLf))
				Else
					Response.write("<td valign=middle width=25 style=" & DblQt("cursor:hand;") & " onClick=" & DblQt("previewPopup('" & ConvertToWebPath(js_EscapePath) & "')") & "><u><font color='#0000FF'>preview</font></u></td>" & vbCrLf)
				End If
			
			Else		
				strImgExt = "jpg,png,bmp,gif"
				If Instr(strImgExt,LCase(GetFSOFileExtension(fil.Name))) < 1 Then
					Response.write("<td valign=middle align='center' width=25>&nbsp;&nbsp;<a href='" & ConvertToWebPath(js_EscapePath) & "'>details</a></td>" & vbCrLf)			
				Else
					Response.write("<td valign=middle width=25 style=" & DblQt("cursor:hand;") & " onClick=" & DblQt("previewPopup('" & ConvertToWebPath(js_EscapePath) & "')") & "><u><font color='#0000FF'>preview</font></u></td>" & vbCrLf)
				End If
			End If
			
			response.Write("</tr>")

		end if

		Call CloseRs(rs)						
		
	Next

	' if there are NO files in the current directory then notify the user
	If blnIsFiles = False Then
		response.write("<div align='center' class='message'>&nbsp;<p>No Files match the filter of<br>(" & strFileExtHeader & ")</p></div>")
	End If
	
	Call CloseDb()

End Function

' PURPOSE: 	Get a list of all the files within a sub-directory
'						based on the filtered criteria.
' INPUTS:		aFolder - this is the folder object to get property values for.
Function GetFSOFiles(aFolder,intType)

'On Error Resume Next

	Dim objLocalFSO
	'----------------------------------------------
	' Database interaction variable declarations
	'----------------------------------------------

	Dim strMode
	Dim strSql
	Dim strFileName
	Dim strFilePath
	Dim blnIsManaged
	Dim intFileId
	Dim intPageId
	Dim strLastModifiedBy
	Dim strIsActive
	Dim strIsLive

	'---------------------------------------------
	' File system variable declarations
	'---------------------------------------------

	Dim fil
	Dim ext
	Dim arrFile
	Dim js_EscapePath
	Dim blnIsFiles
	
	For Each fil In aFolder.files
					

		'-------------------------------------
		' Open recordset object
		'-------------------------------------	
		Call OpenRs(rs)

		' If the current file extension is in the filter list, then we include it.
		If Instr(LCase(strFileExtHeader),LCase(GetFSOFileExtension(fil.Name))) > 0 Then

			strFileName = fil.Name
			strFilePath = ConvertToWebPath(aFolder.Path)
			
			' query database and look for a managed file
			strSql = "SELECT FileId, FileName, FilePath, PageId, FirstName, LastName, f.ActiveFlag, LiveState " & _
					 "FROM tblFiles f, tblUsers u " & _
					 "WHERE f.LastModifiedBy = u.UserId " & _
					 "AND EncryptedFileName = '" & fil.Name & "' " & _
					 "AND FilePath = '" & strFilePath & "'"
			
			'Response.write strSql & "<br>"
			
			rs.Open strSql, GetConnection
			
			' if no records exist, then the file is not managed
			If rs.BOF And rs.EOF Then
				blnIsManaged = False
			Else
				' get values and store in variables
				blnIsManaged = True
				intFileId = rs(0)
				strFileName = rs(1)
				strFilePath = rs(2)
				intPageId = rs(3)
				strLastModifiedBy = rs(4) & " " & rs(5)
				
				If rs(6) Then
					strIsActive = "Active"
				Else
					strIsActive = "Inactive"
				End If
				
				If rs(7) Then
					strIsLive = "LIVE"
				Else
					strIsLive = "STAGE"
				End If
				
			End If
			
			' we set a flag to indicate our message handler that we have records found
			blnIsFiles = True
			
			' if the file is not managed then we want to mark it
			If Not blnIsManaged Then
			
				strMode = "NEW"
				intFileId = 0
				Response.write("<tr onClick='selectRow(this);' onMouseover=this.bgColor='#8DBCEB' onMouseout=this.bgColor='#FFFFFF' style='cursor:hand'>" & vbCRLf)			

			Else

				strMode = "EDIT"
				' we perform some dhtml features on each row
				Response.write("<tr onClick='selectRow(this);' onMouseover=this.bgColor='#8DBCEB' onMouseout=this.bgColor='#FFFFFF' style='cursor:hand'>" & vbCRLf)

			End If

			Response.write("<td valign=top width=150>")
			If Not blnIsManaged Then
				Response.write("<img src='/lambent/cms/static/uma.gif' border='0' align='absmiddle' alt='Unmanaged file'>&nbsp;")
			Else
				Response.write("<img src='/static/icons/" & GetFSOFileExtension(strFileName) & ".gif' border='0' align='absmiddle' alt='" & fil.Name & "'>&nbsp;")
			End If			
			Response.write(strFileName & "</td>" & vbCrLf)
			Response.write("<td valign=middle width=145>" & FormatDate(fil.DateLastModified,"%m/%d/%y - %h:%N %P") & "</td>" & vbCrLf)
			Response.write("<td valign=middle width=130>" & strLastModifiedBy & "</td>" & vbCrLf)

			' get the dimensions of the file. We pass arguments by reference to the gfxSpex function
			' if there are dimensions to the file, then display them, otherwise skip it.
			dim w,h,c,strType
			if gfxSpex(fil.Path, w, h, c, strType) then
				response.write "<td valign=middle width=150>" & w & " x " & h & "</td>"
			else
				response.write "<td width=150>&nbsp;</td>"
			end if
			
			' get the file size
			Response.write("<td valign=middle width=50>" & SizeFormat(fil.size) & "</td>" & vbCrLf)

			Set objLocalFSO = Server.CreateObject("Scripting.FileSystemObject")
			' perform an escape of the backslash to support the path to be passed to javascript functions
			'Response.write objLocalFSO.GetAbsolutePathName(fil.Path) & "<br>"
			
			js_EscapePath = replace(fil.Path,"\","\\")
			'js_EscapePath = rs(2) & "/" & rs(1)
			
			If intType = 1 Then

				If Instr(strImageExt,LCase(GetFSOFileExtension(strFileName))) < 1 Then
					Response.write("<td valign=middle align='center' width=25>&nbsp;&nbsp;<a target='_blank' href='/lambent/cms/files/download.asp?id=" & intFileId & "'>save</a></td>" & vbCrLf)			
				Else
					' format the logic for the "preview" link
					Response.write("<td valign=middle width=25 style=" & DblQt("cursor:hand;") & " onClick=" & DblQt("previewPopup('" & ConvertToWebPath(js_EscapePath) & "')") & "><u><font color='#0000FF'>preview</font></u></td>" & vbCrLf)
				End If
			
			Else		

				' format the logic for the "select" link
				Response.Write("<td valign=middle width=25 style=" & DblQt("cursor:hand;") & " onClick=" & DblQt("selectImage('" & ConvertToWebPath(js_EscapePath) & "','" & fil.Name & "')") & "><u><font color='#0000FF'>select</font></u></td>" & vbCrLf)

				If Instr(strImageExt,LCase(GetFSOFileExtension(fil.Name))) < 1 Then
					Response.write("<td valign=middle align='center' width=25>&nbsp;&nbsp;<a href='/lambent/cms/files/download.asp?id=" & intFileId & "'>save</a></td>" & vbCrLf)			
				Else
					' format the logic for the "preview" link
					Response.write("<td valign=middle width=25 style=" & DblQt("cursor:hand;") & " onClick=" & DblQt("previewPopup('" & ConvertToWebPath(js_EscapePath) & "')") & "><u><font color='#0000FF'>preview</font></u></td>" & vbCrLf)
				End If
				' format the logic for the "del" link
				'Response.write("<td valign=top width=25 style=" & DblQt("cursor:hand;") & " onClick=" & DblQt("deleteImage('" & ConvertToWebPath(js_EscapePath) & "')") & "><u><font color='#0000FF'>delete</font></u></td>" & vbCrLf)		
	'			Response.write("<td valign=top width=25 style=" & DblQt("cursor:hand;") & " onClick=" & DblQt("manageFile('" & ConvertToWebPath(js_EscapePath) & "')") & "><u><font color='#0000FF'>manage</font></u></td>" & vbCrLf)					
				Response.write("<td valign=middle width=25 style=" & DblQt("cursor:hand;") & " onClick=" & DblQt("manageFile(" & intFileId & ",'" & strFilePath & "','" & strFileName & "','" & strMode & "')") & "><u><font color='#0000FF'>manage</font></u></td>" & vbCrLf)					

			End If
			
			response.Write("</tr>")

		end if

		Call CloseRs(rs)						
		
	Next

	' if there are NO files in the current directory then notify the user
	If blnIsFiles = False Then
		response.write("<div align='center' class='message'>&nbsp;<p>No Files match the filter of<br>(" & strFileExtHeader & ")</p></div>")
	End If
	
	Call CloseDb()

End Function

' PURPOSE:	Deletes a file from the system.
' INPUTS:		strFile - file to be deleted.
' NOTES:		We create this global page variable and initialize it to an empty string. We need
' 					to do this because we call this variable within our page to display which file was just 
' 					deleted. If we did not clear this variable out, then the message would remain on the screen
' 					until the window was reloaded.
'Dim strDeletedFile
'strDeletedFile = ""
Function DeleteFSOFile(strFile)
	On Error Resume Next
	Dim fil
	Set fil=objFSO.GetFile(ConvertToFilePath(strFile))
	strDeletedFile = fil.Name
	fil.Delete
	Set fil = Nothing		
	ReloadFSOComboFolder	
End Function

' PURPOSE:	Moves a file from one folder to another.
' INPUTS:		strFile - The file to be moved to another directory
'						strTargetFolder - The target folder for where to move the file
Function MoveFSOFile(strFile,strTargetFolder)
	On Error Resume Next
	dim fil
	If Instr(strFile,strTargetFolder) > 0 then Exit Function
	Set fil = objFSO.GetFile(ConvertToFilePath(strFile))
	fil.Move(strTargetFolder & "\")
	Set fil = Nothing
	ReloadFSOComboFolder
End Function

' PURPOSE:	Copies a file to another directory
' INPUTS:		srcFile - source file that will be copied
'						cpyFile - new file name for the copied srcFile
Function CopyFSOFile(srcFile,cpyFile)
	On Error Resume Next
	Const OverwriteExisting = True
	If Not objFSO.FileExists(srcFile) Then
		'Response.write "convert srcFile = " & ConvertToFilePath(srcFile) & "<br>"
		'Response.write "remove file from path = " & RemoveFileFromPath(ConvertToFilePath(srcFile)) & cpyFile & "<br>"
		objFSO.CopyFile ConvertToFilePath(srcFile), RemoveFileFromPath(ConvertToFilePath(srcFile)) & cpyFile, OverwriteExisting
	Else
		Response.write "<script>alert('A file name already exists with the file name.\nPlease select a different name.');</script>"
	End If
	ReloadFSOComboFolder
End Function

' PURPOSE:	Rename a file in the filesystem
' INPUTS:		srcFile - The file that will be renamed
'						renFile - The new file name
Function RenameFSOFile(srcFile,renFile)
	On Error Resume Next
	If Not objFSO.FileExists(renFile) Then
		objFSO.MoveFile ConvertToFilePath(srcFile), renFile
	Else
		Response.write "<script>alert('The file name to which you are renaming already exists.');</script>"
	End If
	ReloadFSOComboFolder
End Function

' PURPOSE:	To get the filename from a fullpath
'	INPUTS:		strDirFolder - Fullpath of the directory including the filename at end
function ParseFileFromPath(strDirFolder)
	Dim pathPos
	pathPos = InStrRev(strDirFolder, "\")
	If pathPos = 0 Or IsNull(pathPos) Then
		ParseFileFromPath = strDirFolder
		Exit Function
	End If
	ParseFileFromPath = Right(strDirFolder, Len(strDirFolder) - pathPos)
end function

' PURPOSE:	To get only the path to the current file. 
'	INPUTS:		strFullPath - Fullpath, including the filename, in which we want to parse
Function RemoveFileFromPath(strFullPath)
	Dim pathPos
	pathPos = InStrRev(strFullPath, "\")
	If pathPos = 0 Or IsNull(pathPos) Then
		RemoveFileFromPath = strFullPath
		Exit Function
	End If
	RemoveFileFromPath = Left(strFullPath,pathPos)
End Function

' PURPOSE:	Formats the current path to be the new pathname. If the srcFolder is 
'						/static/images/products/bicycles and we want to change bicycles to bikes 
'						then this routine will concatenate /static/images/products/ & bikes together.
' INPUTS:		srcFolder - source folder that we need to parse
'						strNewDirName - new folder name that we need to append to the parsed srcFolder
Function ChangeDirName(srcFolder,strNewDirName)
	Dim pathPos
	pathPos = InStrRev(srcFolder, "\")
	If pathPos = 0 Or IsNull(pathPos) Then
		ChangeDirName = srcFolder
		Exit Function
	End If
	ChangeDirName = Left(srcFolder,pathPos) & strNewDirName
End Function

' PURPOSE: 	Renames a folder	
'	INPUTS:		srcFolder - Original folder that will be renamed
'						strNewFolder - The new folder name
Function RenameFSOFolder(srcFolder,strNewFolder)
	On Error Resume Next
	strNewFolder = ChangeDirName(srcFolder,strNewFolder)
	If Not FolderExists(strNewFolder) Then
		objFSO.MoveFolder srcFolder, strNewFolder
	Else
		Response.write "<script>alert('The directory already exists. Please choose a different name.');</script>"
	End If
	strTargetFolder = strNewFolder
	ReloadFSOComboFolder
End Function

' PURPOSE:	Creates a new directory	
'	INPUTS:		strNewFolder - new /path/folder name to be created
Function NewFSOFolder(strNewFolder)
	On Error Resume Next
	' we are only allowing folders to be created 3 levels deep
	If CharCount(GetFormPost("selFolderList"),"\") - CharCount(strComboTopLevel,"\") > 2 Then
		Response.write "<script>alert('You are not allowed to create directories more than three (3) \n" & _
									"levels deep under the root directory in this application.');</script>"
	Else
		If Not FolderExists(strNewFolder) Then
			Set objFolder = objFSO.CreateFolder(strNewFolder)
		Else
			Response.write "<script>alert('The directory already exists. Please choose a different name.');</script>"		
		End If
	End If
		ReloadFSOComboFolder		
End Function

' PURPOSE:	Delets a particular folder from the file system	
'	INPUTS:		strFolder - directory to be deleted.
Dim strTargetFolder
Function DeleteFSOFolder(strFolder)
	On Error Resume Next
	objFSO.DeleteFolder(strFolder)
	strTargetFolder = strComboTopLevel
	ReloadFSOComboFolder	
End Function

' PURPOSE:	Moves a folder from one directory to another
'	INPUTS:		srcFolder - The directory to be moved
'						targetFolder - The parent directory for which new folder will be a child of
Function MoveFSOFolder(srcFolder,targetFolder)
	On Error Resume Next
	If FolderCheck(srcFolder,targetFolder) Then
		Set objFolder = objFSO.GetFolder(srcFolder)
		objFolder.Move(targetFolder & "\")
		strTargetFolder = targetFolder
	End If
	ReloadFSOComboFolder
End Function



Function GetFileExtension(strFile)

	Dim dotPos

	dotPos = InstrRev(strFile,".")
	GetFileExtension = Right(strFile,(Len(strFile)-dotPos))

End Function


Function GetFileName(strPath)
	
	GetFileName = Right(strPath,Len(strPath)-InstrRev(strPath,"\"))

End Function

function GetFilePath(strDirFolder)
	Dim pathPos
	pathPos = Instr(strDirFolder, "/")		
	If pathPos = 0 Or IsNull(pathPos) Then
		GetFilePath = strDirFolder
		Exit Function
	End If
	GetFilePath = Left(strDirFolder, Len(strDirFolder) - Len(GetFileName(replace(strDirFolder,"/","\"))))
end function

' Function: DownloadFile
' Purpose: 	Downloads a file from the website
' Inputs:	strPath = Web server path to the file on the file system
'			strFile = File to download
' Outputs: 	The "Save As" dialog box to download the file.
' Comments: This function works against secure/non-secure directories

Function DownloadFile(strPath,strFile)
	
	Dim adoStream

    Response.ContentType = "application/asp-unknown" ' arbitrary 
    Response.AddHeader "content-disposition","attachment; filename=" & strFile
    Set adoStream = Server.CreateObject("ADODB.Stream") 

	With adoStream
    	.Open() 
		.Type = 1 
    	.LoadFromFile(ConvertToFilePath(strPath) & "\" & strFile)     
		Response.BinaryWrite .Read() 
    	.Close()
	End With
	
	Set adoStream = Nothing 	

	Response.End()
	
End Function
%>
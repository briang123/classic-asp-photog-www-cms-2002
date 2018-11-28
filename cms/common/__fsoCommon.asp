<%
'********************************************************************************
' START COMMON FILE FUNCTIONS
'********************************************************************************
'
' MODULE: 		fsoCommon.asp
' CREATED BY: 	Brian Gaines
' CREATED ON: 	02/27/2003
'
' NOTE: 		The functions GetBytes, lngConvert, lngConvert2, gfxSpex were 
'				not created by me. They were picked up on the web and developed 
'				by someone else, but I don't recall his name.  He did mention 
'				in his copyright to be able to use the code freely.
'				
'				All other code in this module I created.
'
'********************************************************************************
' START IMAGE DIMENSION FUNCTIONS
'********************************************************************************

' PURPOSE: 	This routine will attempt to identify any filespec passed 
'			as a graphic file (regardless of the extension). This will			
'			work with BMP, GIF, JPG and PNG files. This function gets 
'			a specified number of bytes from any file, starting at the 
'			offset (base 1)
' INPUTS:	flnm - Filespec of file to read
'			offset - Offset at which to start reading
'			bytes - How many bytes to read
function GetBytes(flnm, offset, bytes)

Dim objFSO, objFTemp, objTextStream, lngSize, strBuff

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
'			to read any file, regardless of the extension, and will
'			identify if it is a graphical image.
' INPUTS: 	flnm - Filespec of file to read
'			width - width of image   
'			height - height of image  
'			depth - color depth (in number of colors)
'			strImageType - type of image (e.g. GIF, BMP, etc.)
function gfxSpex(flnm, width, height, depth, strImageType)

	dim strPNG, strGIF, strBMP, strType, strBuff, lngSize, flgFound, strTarget
	dim ExitLoop, lngPos

	strType = ""
	strImageType = "(unknown)"
	
	gfxSpex = False
	
	strPNG = chr(137) & chr(80) & chr(78)
	strGIF = "GIF"
	strBMP = chr(66) & chr(77)
	
	strType = GetBytes(flnm, 0, 3)
	
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

function ImageResize(strImageName, intDesiredWidth, intDesiredHeight)
'http://www.4guysfromrolla.com/webtech/011201-1.shtml
  dim TargetRatio
  dim CurrentRatio
  dim strResize
  dim w, h, c, strType

  if gfxSpex(strImageName, w, h, c, strType) = true then
	 TargetRatio = intDesiredWidth / intDesiredHeight
	 CurrentRatio = w / h
	 if CurrentRatio > TargetRatio then                       ' We'll scale height
		strResize = "width=""" & intDesiredWidth & """"
	 else
		strResize = "height=""" & intDesiredHeight & """"     ' We'll scale width
	 end if
  else
	 strResize = ""
  end if

  ImageResize = strResize

end Function


'********************************************************************************
' START FILE SYSTEM OBJECT SUPPORT FUNCTIONS
'********************************************************************************

Sub ReloadFSOComboFolder
	Set objFolder = objFSO.GetFolder(strComboTopLevel)
End Sub

' PURPOSE: 	This function formats the file size into a value that
'			is a bit more readable. The number is read in as bytes.	
' INPUTS:	number - This is the file size
Function SizeFormat(number)
	If number < 1000 Then
		SizeFormat = number & " Bytes"
	ElseIf number > 999 And number < 1000000 Then
		number = Round(number/1000)
		SizeFormat = number & " Kb"
	ElseIf number > 1000000 Then
		number = Round(number/1000000)
		SizeFormat = number & " Mb"
	End If
End Function

' PURPOSE: 	Converts a path formatted as "C:\\Inetpub\\wwwroot\\..."
'			into a web server path such as "/appfolder/dir1/file1.gif"
' INPUTS:	strPath - Path to where the file is stored.	
Function ConvertToWebPath(strPath)

	Dim normalized, stripRoot
	' reformat the javascript escaping slashes
	normalized = replace(strPath,"\\","\")
	
	' now that we have the format of C:\Inetpub\wwwroot\nlt\...
	' we can hide the non-web browser path. where:
	' strRootPath is a user-defined variable that defines the application directory
	' strBaseWebPath is a user-defined variable that represents the base web application path
	stripRoot = replace(normalized,strRootPath,strBaseWebPath) 
	
	' One last formatting technique before we return this value to the caller.
	' change all the backward slashes to forward slashes to be referenced as web server path.
	ConvertToWebPath = replace(stripRoot,"\","/")
	
End Function

Function GetFilePath(strPath)
	GetFilePath = strRootWebServer & replace(strPath,"/","\")
End Function

' PURPOSE: 	Converts a web server path formatted as "/NLT/static/images/"
'						to "C:\Inetpub\wwwroot\NLT\static\images\file1.gif"
' INPUTS: 	strPath - Path to convert
Function ConvertToFilePath(strPath)
'	ConvertToFilePath = strRootWebServer & replace(strPath,"/","\")
	ConvertToFilePath = replace(strPath,"/","\")
End Function

' PURPOSE:	Get the file extension of the file.
' INPUTS:		strFile - The file
Function GetFSOFileExtension(strFile)
	
	Dim arrFile, ext
	
	' get the file extension. 
	If Instr(strFile,".") > 0 then
		arrFile = split(strFile,".")
		GetFSOFileExtension = arrFile(1)
	else ' return some bogus extension - i think the tilda will do
		GetFSOFileExtension = "~~~"
	end if
		
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
		valToMatch = strRootPath
	end if

	Dim strDisplayPath
	If fsoFolder = strComboTopLevel Then
		strDisplayPath = "\"
	Else
		strDisplayPath = Replace(fsoFolder.ParentFolder,strComboTopLevel,"") & "\" & fsoFolder.Name
	End If
	Response.Write("<OPTION VALUE=" & DblQt(fsoFolder) & setComboOption(fsoFolder,valToMatch,1) & ">" & strDisplayPath & "</OPTION>" & vbCRLF)
					
	' Loop through each sub-folder of the current folder object
	Dim subFolder
	For Each subFolder in fsoFolder.SubFolders
	  	Call BuildFSOFolderCombo(subFolder,valToMatch)
	Next	

End Function

' PURPOSE: 	Get a list of all the files within a sub-directory
'						based on the filtered criteria.
' INPUTS:		aFolder - this is the folder object to get property values for.
Function GetFSOFiles(aFolder)

	Dim fil,ext,arrFile,js_EscapePath,blnIsFiles

	For Each fil In aFolder.files
					
		' If the current file extension is in the filter list, then we include it.
		If Instr(LCase(strFileExtHeader),LCase(GetFSOFileExtension(fil.Name))) > 0 Then
			
			' we set a flag to indicate our message handler that we have records found
			blnIsFiles = True
			
			' we perform some dhtml features on each row
			Response.write("<tr onClick='selectRow(this);' onMouseover=this.bgColor='#8DBCEB' onMouseout=this.bgColor='#FFFFFF' style='cursor:hand'>" & vbCRLf)

			' get the file name
			Response.write("<td valign=top>" & fil.Name & "</td>" & vbCrLf)

			' get the dimensions of the file. We pass arguments by reference to the gfxSpex function
			' if there are dimensions to the file, then display them, otherwise skip it.
			dim w,h,c,strType
			if gfxSpex(fil.Path, w, h, c, strType) then
				response.write "<td valign=top>" & w & " x " & h & "</td>"
			else
				response.write "<td>&nbsp;</td>"
			end if
			
			' get the file size
			Response.write("<td valign=top>" & SizeFormat(fil.size) & "</td>" & vbCrLf)

			' perform an escape of the backslash to support the path to be passed to javascript functions
			js_EscapePath = replace(fil.Path,"\","\\")

			' format the logic for the "select" link
			Response.Write("<td valign=top style=" & DblQt("cursor:hand;") & " onClick=" & DblQt("selectImage('" & ConvertToWebPath(js_EscapePath) & "','" & w & "','" & h & "','" & fil.Name & "')") & "><u><font color='#0000FF'>select</font></u></td>" & vbCrLf)

			' format the logic for the "del" link
			Response.write("<td valign=top style=" & DblQt("cursor:hand;") & " onClick=" & DblQt("deleteImage('" & ConvertToWebPath(js_EscapePath) & "')") & "><u><font color='#0000FF'>del</font></u></td>" & vbCrLf)		
			response.Write "</tr>"
		end if
	Next

	' if there are NO files in the current directory then notify the user
	If blnIsFiles = False Then
		response.write("<div align='center' class='message'>&nbsp;<p>No Files match the filter of<br>(" & strFileExtHeader & ")</p></div>")
	End If
	
End Function

' PURPOSE: 	Deletes a file from the system.
' INPUTS:		strFile - file to be deleted.
' NOTES: 		We create this global page variable and initialize it to an empty string. We need
' 					to do this because we call this variable within our page to display which file was just 
' 					deleted. If we did not clear this variable out, then the message would remain on the screen
' 					until the window was reloaded.
Dim strDeletedFile
strDeletedFile = ""
Function DeleteFSOFile(strFile)

	Dim fil

	' convert the current file format to one that is recognizable by the FSO.
	If instr(lcase(strFile),"c:") > -1 Then
	    Set fil=objFSO.GetFile(strFile)
	Else
	    Set fil=objFSO.GetFile(Server.MapPath(ConvertToFilePath(strFile)))
	End If
	
	' we now set the strDeletedFile variable so it can be displayed to the user.
	strDeletedFile = fil.Name
	
	' we delete the file from the system.
	fil.Delete
	
	' destroy the file object
	Set fil = Nothing		
	
End Function

' PURPOSE:	Moves a file from one folder to another.
' INPUTS:		strFile - The file to be moved to another directory
'						strTargetFolder - The target folder for where to move the file
Function MoveFSOFile(strFile,strTargetFolder)

	' we check to see if the source folder and target folder are the same.
	' if the folders are the same we exit the function and do nothing.
	If Instr(strFile,strTargetFolder) > 0 then
		Exit Function
	End If
	
	dim fil

	' convert the current file format to one that is recognizable by the FSO.
	Set fil = objFSO.GetFile(ConvertToFilePath(strFile))

	' move the file to the target folder
	fil.Move(strTargetFolder & "\")
	
	' destroy the file object
	Set fil = Nothing

End Function

' PURPOSE:	Copies a file to another directory
' INPUTS:		srcFile - 
'						cpyFile - 
Function CopyFSOFile(srcFile,cpyFile)

	response.Write(objFSO.FileExists(srcFile) & " :: " & cpyFile)

'	If Not objFSO.FileExists(cpyFile) Then
	
	
		
'		Dim fil
		' convert the current file format to one that is recognizable by the FSO.
'		Set fil = objFSO.GetFile(ConvertToFilePath(srcFile))
		' copy the file to the same directory with the new name
'		fil.Copy fil.ParentFolder & "\" & cpyFile,False
		' destroy the file object
'		Set fil = Nothing
'	Else
'		Response.write "<script>alert('A file name already exists with the file name.\nPlease select a different name.');</scrip>
'	End If
'	ReloadFSOComboFolder
		
End Function


Function RenameFSOFile(srcFile,renFile)
	If Not objFSO.FileExists(renFile) Then
		objFSO.MoveFile ConvertToFilePath(srcFile) , renFile
	Else
		Response.write "<script>alert('The file name to which you are renaming already exists.');</script>"
	End If
	ReloadFSOComboFolder
		
End Function

Function RenameFSOFolder(srcFolder,strNewFolder)


End Function

Function NewFSOFolder(strNewFolder)
	Set objFolder = objFSO.CreateFolder(strNewFolder)
	ReloadFSOComboFolder	
End Function

Function DeleteFSOFolder(strFolder)
	objFSO.DeleteFolder(strFolder)
	ReloadFSOComboFolder	
End Function

Function MoveFSOFolder(srcFolder,targetFolder)

	If FolderCheck(srcFolder,targetFolder) Then
		Set objFolder = objFSO.GetFolder(srcFolder)
		objFolder.Move(targetFolder & "\")
	End If
	ReloadFSOComboFolder

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
		Response.write "<script>alert('You are no able to perform this operation because the source and \n" & _
										"target directory are pointing to the same directory.');</script>"
	End If
End Function

%>
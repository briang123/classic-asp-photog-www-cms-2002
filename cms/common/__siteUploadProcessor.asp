<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/cms/common/__fsoConfig.asp" -->
<!-- #include virtual="/cms/common/__fsoCommon.asp" -->
<!-- #include virtual="/objects/cFileUpload.asp"-->
<!-- #include virtual="/objects/cPhotos.asp" -->
<%
'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function AddPhotosToDb(lgImg,pageId,imgOrder,lngPhotoId)

	Dim oPhoto
	Set oPhoto = New cPhotos
	With oPhoto
		.WebPageId = pageId
		.LargeImage = lgImg
		.ImageOrder = imgOrder
		.ActiveFlag = 0
		.AddSitePhoto()
		lngPhotoId = .ID
		AddPhotosToDb = Not .IsError
	End With
	Set oPhoto = Nothing

End Function

Function DeletePhoto(id,largeFile)

	If StringNotEmptyOrNull(largeFile) Then
		Call DeleteFSOFile(largeFile)
	End If
	
	Dim oPhoto
	Set oPhoto = New cPhotos
	With oPhoto
		.ID = id
		.DeleteSitePhotos()
		DeletePhoto = Not .IsError
	End With
	
End Function

Dim blnSuccess,lngPhotoId,uploadCount
uploadCount = 0

Dim invalidFile
Dim path
Dim intPageId

path = GetQryString("p")

Dim Upload, Count, ErrorNum, Descr
Set Upload = Server.CreateObject("SoftArtisans.FileUp")

Dim LargeImagePath
Dim imgPath
imgPath = GetAppVariable("IMAGE_PATH")
LargeImagePath = "\" & Replace(imgPath,"/","\")

Upload.Path = Upload.Form("hidUploadPath") & LargeImagePath
'echobr(Upload.Path)

Upload.PreserveMacBinary = True
Upload.MaxBytes = MAX_FILE_KB_UPLOAD_SIZE * 1000 '200 kb
Upload.MaxBytesToCancel = 4000000 '4 mb

'On Error Resume Next
ErrorNum = Err.Number
Descr = Err.Description
'On Error Goto 0

intPageid = GetQryString("fid") 'Upload.Form("hidPageId")

'echobr(intPageId)

invalidFile = ""	
If ErrorNum <> 0 Then
	invalidFile = "The following error occurred: " & ErrorNum & " " & Descr
	Upload.Flush
Else
dim temp
	For Each element in Upload.FormEx
        If lcase(left(element,4))="file" then
		    If IsObject(Upload.Form(element)) Then
                If Not Upload.Form(element).IsEmpty Then
	                Dim allowed,i,isAllowed
	                allowed = Split(FEXT_ALLOWED,";")
	                isAllowed = False

	                Dim FName, FExtension, FCONT
	                'FName = Mid(Upload.UserFilename, InstrRev(Upload.UserFilename, "\") + 1)			            
	                FName = Mid(Upload.Form(element).UserFilename, InstrRev(Upload.Form(element).UserFilename, "\") + 1)
                    FExtension = MID(FName,instr(FName,".")+1)
                    For i = LBound(allowed) To UBound(allowed)
                        isAllowed = (lcase(allowed(i))=lcase(FExtension))
		                If isAllowed Then
		                    Exit For
		                End If
	                Next
                End If 
            End If
			
		    If isAllowed = False Then
			    invalidFile = invalidFile & "\n\n-" & FName & " (Exception: Invalid file format. Allowable types are (" & UCase(FEXT_ALLOWED) & "))"
		    End If			
			
            If IsObject(Upload.Form(element)) Then
                If Not Upload.Form(element).IsEmpty Then
'echobr(element)
'echobr(Upload.Form(element).UserFileName)
                    'On Error Resume Next
                        Upload.Form(element).Save
'echobr(err.number)
                        If Err.Number <> 0 Then                              
                            ErrorNum = Err.Number
                            Descr = Err.Description
                            invalidFile = invalidFile & "\n\n-An error occurred while saving the file: " & FName & ": " & ErrorNum & ": " & Descr
                        Else            
                            uploadCount = uploadCount + 1
'echobr(fname)
                            blnSuccess = AddPhotosToDb(FName,intPageId,uploadCount,lngPhotoId)  
'echobr(blnSuccess)
          
	                        If blnSuccess = False Then
				                invalidFile = invalidFile & "\n\n-" & FName & " (Exception: Database Error)"
				                blnSuccess = DeletePhoto(lngPhotoId,IMAGE_PATH & "/" & FName)
                            End If                                
                        End If
                    'On Error Goto 0
                End If 
            End If
        End If
	Next
End If	

If invalidFile <> "" Then
    Upload.Flush
	Response.Write("<script>alert('" & invalidFile & "\n\nPlease correct any problems and re-upload those files.');</script>")
	Response.Write("<script>window.location.href='__siteUploadForm.asp?msg=fail';</script>")
    Set Upload = Nothing
Else
    Upload.Flush
    Set Upload = Nothing
'echobr("here")
	'Response.Write("<script>window.location.href='__siteUploadForm.asp?msg=success&fid=" & intPageid & "&uloc=site';</script>")
	Response.Redirect("__siteUploadForm.asp?msg=success")	
End If


%>

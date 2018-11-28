<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/cms/common/__fsoConfig.asp" -->
<!-- #include virtual="/cms/common/__fsoCommon.asp" -->
<!-- #include virtual="/objects/cFileUpload.asp"-->
<%

'Dim debug
Sub AppendDebugInfo(key,value)

	debug = debug & UCASE(key) & ": " & value & "<BR>"

End Sub


	Dim invalidFile
	Dim strGalleryLastName,strCategoryName
	Dim path
	
	path = GetQryString("p")

	'Call AppendDebugInfo("querystring path",path)


	Dim Upload, Count, ErrorNum, Descr
	Set Upload = Server.CreateObject("SoftArtisans.FileUp")
	Upload.Path = Upload.Form("hidUploadPath") & path
    	'Upload.PreserveMacBinary = True
	Upload.MaxBytes = 20000 '200 kb
	Upload.MaxBytesToCancel = 4000000 '4 mb
	
	'On Error Resume Next
	ErrorNum = Err.Number
	Descr = Err.Description
	'On Error Goto 0

	strGalleryLastName = Upload.Form("hidGalleryLastName")
	strCategoryName = Upload.Form("hidCategoryName")

	'Call AppendDebugInfo("gallery last name",strGalleryLastName)
	'Call AppendDebugInfo("category name",strCategoryName)


	invalidFile = ""	
	If ErrorNum <> 0 Then
		invalidFile = "The following error occurred: " & ErrorNum & " " & Descr
		Upload.Flush
	Else
		For Each element in Upload.FormEx

'Upload.form(element).PreserveMacBinary = true

'response.write(element)
'response.end

'Upload(element).PreserveMacBinary = True
	        If lcase(left(element,4))="file" then
			    If IsObject(Upload.Form(element)) Then
	                If Not Upload.Form(element).IsEmpty Then
		                Dim allowed,i,isAllowed
		                allowed = Split(FEXT_ALLOWED,";")
		                isAllowed = False
    		            
	'Call AppendDebugInfo("file name",Upload.Form(element).UserFilename)

		                Dim FName, FExtension, FCONT
		                FName = Mid(Upload.Form(element).UserFilename, InstrRev(Upload.Form(element).UserFilename, "\") + 1)

	'Call AppendDebugInfo("FName",FName)

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
	                    On Error Resume Next
		                    Upload.Form(element).Save
                            If Err.Number <> 0 Then                              
	                            ErrorNum = Err.Number
	                            Descr = Err.Description	                        
	                            invalidFile = invalidFile & "\n\n-An error occurred while saving the file: " & FName & ": " & ErrorNum & ": " & Descr
                            End If
 		                On Error Goto 0
	                End If 
                End If
            End If
		Next
	End If	

	if invalidFile <> "" then
	
	'Call AppendDebugInfo("error info",invalidFile)

	'SendDebugInfo

	    Upload.Flush
		response.write("<script>alert('" & invalidFile & "');</script>")
		If Len(strGalleryLastName) > 0 Then
			Response.Write("<script>window.location.href='__uploadForm.asp?gln=" & strGalleryLastName & "&msg=fail';</script>")
		ElseIf Len(strCategoryName) > 0 Then
			Response.Write("<script>window.location.href='__uploadForm.asp?cn=" & strCategoryName & "&msg=fail';</script>")
		End If
	else
	'SendDebugInfo

		If Len(strGalleryLastName) > 0 Then
			Response.Redirect("__uploadForm.asp?gln=" & strGalleryLastName & "&msg=success")
		ElseIf Len(strCategoryName) > 0 Then
			Response.Redirect("__uploadForm.asp?cn=" & strCategoryName & "&msg=success")
		End If			
	end if
	
	Set Upload = Nothing


Sub SendDebugInfo
        Dim mailer
        Set mailer = CreateObject("SoftArtisans.SMTPMail")
        With mailer
            	.RemoteHost = "mail.juliestarkphotography.com"
            	.Subject = "Email from JulieStarkPhotography.com - " & FormatDateTime(Now(),1)
            	.HtmlText = debug
            	.AddRecipient "", "bgaines@newleaftechinc.com"
		'.AddRecipient "", "sohophotography@yahoo.com"
            	.FromAddress = "website@juliestarkphotography.com"
		.ReplyTo  = "julie@juliestarkphotography.com"
		.UserName= "website@juliestarkphotography.com"
		.Password= "website"
             	.SendMail()
        End With
End Sub


%>

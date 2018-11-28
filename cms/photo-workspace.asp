<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/cms/common/__fsoConfig.asp" -->
<!-- #include virtual="/cms/common/__fsoCommon.asp" -->
<!-- #include virtual="/objects/cPhotos.asp" -->
<%
Dim strPhotoType
Dim strFolder
Dim strFsoFolder
Dim PHOTO_IMAGE_PATH
Dim QSTRING 
Dim galleryId
Dim strSubTitle
Dim PHOTO_TYPE_PHYSICAL_THUMB_PATH
Dim PHOTO_TYPE_PHYSICAL_LARGE_PATH

If StringNotEmptyOrNull(GetQryString("gid")) Then
	strSubTitle = GetQryString("gn")
	strPhotoType = "GALLERY"
	QSTRING = "?gid=" & GetQryString("gid") & "&gln=" & GetQryString("gln") & "&gn=" & GetQryString("gn")
	galleryId = GetQryString("gid")
	PHOTO_IMAGE_PATH = ROOT_PATH & PROOF_PATH & "/" & GetQryString("gln")
	''''''''''PHOTO_IMAGE_PATH = replace("c:\Websites" & PHOTO_IMAGE_PATH & "\thumbs\" & GetQryString("delete"),"/","\")
	
	'uncomment for localhost development
	'strFsoFolder = "../" & PROOF_PATH & "/" & GetQryString("gln")
	
	'uncomment for production server
	strFsoFolder = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & PROOF_PATH & "\" & GetQryString("gln"),"/","\")
	
	'strFsoFolder = "c:\starktemp\secure\julie\larson"
    '''''''''''die(strFsoFolder)
Else
	strSubTitle = GetQryString("cn")
	strPhotoType = "CATEGORY"
	QSTRING = "?cid=" & GetQryString("cid") & "&cn=" & GetQryString("cn")
	categoryId = GetQryString("cid")
	PHOTO_IMAGE_PATH = ROOT_PATH & GALLERY_PATH & "/" & GetQryString("cn") 'Replace(GetQryString("cn")," ","")

    'uncomment for localhost development
	'strFsoFolder = "../" & GALLERY_PATH & "/" & GetQryString("cn")
	'strFsoFolder = "c:\starktemp\secure\julie\larson\thumbs"

    'uncomment for production server
    strFsoFolder = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & GALLERY_PATH & "\" & GetQryString("cn"),"/","\")
	
    ''''''''die(strFsoFolder)	
End If

'echobr(strFsoFolder)
'''Response.ContentType = "image/jpeg"
'''Set Download = Server.CreateObject("SoftArtisans.FileUp")
'''Download.TransferFile strFsoFolder & "\ma29.jpg"


PHOTO_TYPE_PHYSICAL_THUMB_PATH = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "\" & PHOTO_IMAGE_PATH & "\thumbs\","/","\")
PHOTO_TYPE_PHYSICAL_LARGE_PATH = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "\" & PHOTO_IMAGE_PATH & "\","/","\")

PAGE_TITLE = "Photograph Workspace &lt; " & strSubTitle & " &gt;"

'die(PHOTO_IMAGE_PATH)
Function SizeFormat(number)
	If number < 1000 Then
		SizeFormat = number & " Bytes"
	ElseIf number > 999 And number < 1000000 Then
		number = Round(number/1000)
		SizeFormat = number & " Kb"
	ElseIf number > 1000000 Then
		number = Round(number/1000000,2)
		SizeFormat = number & " Mb"
	End If
End Function

Sub AddToArray(ByRef arr, newItem)
	Redim Preserve arr(Ubound(arr) + 1)
	arr(UBound(arr)) = Trim(newItem)
End Sub

'determine if element is part of array
Function Contains(arr,ByVal imgName)
	Dim el
	For el = LBound(arr) to UBound(arr) 
		If StrComp(arr(el),Trim(imgName),1) = 0 Then
			Contains = True
			Exit Function
		End if
	Next
	Contains = False
End Function

Function DeleteFile(fileName)
	If StringNotEmptyOrNull(fileName) Then
		Call DeleteFSOFile(fileName)
	End If
End Function
	
'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function ConnectPhotos(ptype,id,lgImg,thumbImg,caption,imgOrder,lngPhotoId)

	Dim oPhoto
	Set oPhoto = New cPhotos
	With oPhoto
		.LargeImage = lgImg
		.ThumbImage = thumbImg
		.Caption = caption
		.ImageOrder = imgOrder
		.ActiveFlag = 0
		if UCase(ptype) = "GALLERY" Then
			.GalleryId = id
			.AddGalleryPhoto()
		ElseIf UCase(ptype) = "CATEGORY" Then
			.CategoryId = id
			.AddCategoryPhoto()
		End If
		lngPhotoId = .ID
		ConnectPhotos = Not .IsError
	End With
	Set oPhoto = Nothing

End Function

'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim lngPhotoId
Dim lgImg
Dim thumbImg
Dim errorList
'Dim galleryId
errorList = ""

'----------------------------------------------------------------------------------------
' PAGE RENDER LOGIC
'----------------------------------------------------------------------------------------
If StringNotEmptyOrNull(GetQryString("delete")) And StringNotEmptyOrNull(GetQryString("dtype")) Then
	Select Case GetQryString("dtype")
		Case "large"
		    'uncomment production server
			'path = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & GetAppVariable("GALLERY_PATH") & "\" & GetQryString("cn") & "\" & GetQryString("delete"),"/","\")
			
			path = strFsoFolder & "\" & GetQryString("delete")

			'uncomment localhost development
			'path = replace(PHOTO_TYPE_PHYSICAL_LARGE_PATH & GetQryString("delete"),"/","\")

			Call DeleteFile(path)
		Case "thumb"
			'uncomment production server
			'tpath = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & GetAppVariable("GALLERY_PATH") & "\" & GetQryString("cn") & "\thumbs\" & GetQryString("delete"),"/","\")	
			
			tpath = strFsoFolder & "\thumbs\" & GetQryString("delete")

			'uncomment localhost development
			'tpath = replace(PHOTO_TYPE_PHYSICAL_THUMB_PATH & GetQryString("delete"),"/","\")

			Call DeleteFile(tpath)
	End Select
	PageRedirect("photorefresh.asp" & QSTRING & "&ptype=" & strPhotoType)
End If

If GetFormPost("hidConnect") = "true" Then
	lngPhotoId = GetFormPost("hidPhotoId")
	galleryId = GetFormPost("hidGalleryId")
	categoryId = GetFormPost("hidCategoryId")
	prevImageOrder = CInt(GetFormPost("hidPrevImageOrder"))

	Dim frmCnxn, arrImg, arrFrm
	frmCnxn = GetFormPost("selCnxn")
	arrFrm = Split(frmCnxn,",")
	
	Dim imageOrder
	For i = LBound(arrFrm) to UBound(arrFrm)
		imageOrder = prevImageOrder + i + 1
		arrImg = Split(arrFrm(i)," <==> ")
		lgImg = Trim(arrImg(0)) & ".jpg"
		thumbImg = Trim(arrImg(1)) & ".jpg"

        'die(strPhotoType & "<br>" & Eval(LCase(strPhotoType)&"Id") & "<br>" & lgImg & "<br>" & thumbImg & "<br>" & imageOrder & "<br>" & lngPhotoId)
		'The line Eval(LCase(strPhotoType)&"Id") = the value stored in the variable galleryId or categoryId
		blnSuccess = ConnectPhotos(strPhotoType,Eval(LCase(strPhotoType)&"Id"),lgImg,thumbImg,"",imageOrder,lngPhotoId)
		If blnSuccess = False Then
			errorList = errorList & lgImg & " <==> " & thumbImg & "<br>"
		End If
	Next

	If blnSuccess Then
		displayMessage = "The photographs were successfully related to one another."
		PageRedirect("photorefresh.asp" & QSTRING & "&ptype=" & strPhotoType)
	Else
		displayMessage = "An error occurred while trying to save information to the database. The following connections were not saved: <br>" & errorList
	End If
End If

Dim largePhotoGallery
Dim thumbPhotoGallery
Dim prevImageOrder
prevImageOrder = 0
largePhotoGallery = Array()
thumbPhotoGallery = Array()
Set oPhoto = New cPhotos
Set collPhotos = New cPhotos

If strPhotoType = "GALLERY" Then
	collPhotos.GalleryId = galleryId
	Call collPhotos.GetPhotosByGallery()
ElseIf strPhotoType = "CATEGORY" Then
	collPhotos.CategoryId = categoryId
	Call collPhotos.GetPhotosByCategory()
End If

For Each oPhoto In collPhotos.Photos.Items	
'    echo(oPhoto.LargeImage & "::" & oPhoto.ThumbImage & "<br>")
	Call AddToArray(largePhotoGallery,oPhoto.LargeImage)
	Call AddToArray(thumbPhotoGallery,oPhoto.ThumbImage)
	prevImageOrder = prevImageOrder + 1
	Set oPhoto = Nothing
Next
Set collPhotos = Nothing
%>
<HTML>
<HEAD>
<TITLE><%=PAGE_TITLE%></TITLE>
<script src="/cms/lib/__domradio.js" type="text/javascript"></script>
<!--#include virtual="/cms/styles.asp"-->
</HEAD>
<BODY>
<form name="form1" id="form1" method="post" action="photo-workspace.asp<%=QSTRING%>">
<input type="hidden" id="hidConnect" name="hidConnect" value="">
<input type="hidden" id="hidPhotoId" name="hidPhotoId" value="<%=lngPhotoId%>">
<input type="hidden" id="hidGalleryId" name="hidGalleryId" value="<%=galleryId%>">
<input type="hidden" id="hidCategoryId" name="hidCategoryId" value="<%=categoryId%>">
<input type="hidden" id="hidPrevImageOrder" name="hidPrevImageOrder" value="<%=prevImageOrder%>">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" style="border-bottom-style:solid;border-width:1;border-color:#666;">
	<!-- #include virtual="/cms/common/__header.asp"-->
	<!-- #include virtual="/cms/common/__titlebar.asp" -->
	<tr>
		<td width="100%">
			<table width="100%" border="0" cellpadding="0" cellspacing="0" height="100%">
				<tr>
					<td id="leftnav">
						<!-- #include virtual="/cms/sidenav.asp" -->
					</td>
					<td id="mainbody">		
						<!-- START PAGE BODY TOOLBAR -->	
						<div style="border-bottom:1px solid <%=abelard_border_color%>;width:100%;padding-bottom:20px;">
							<div style="align:left;width:<%=CMS_PAGE_WIDTH%>;">	
								<span style="float:left;padding-right:10px;">&nbsp;</span>
								<span style="float:right;padding-right:10px;">
									<A href="#" onClick="javascript:checkForUnsubmittedConnections();" class="menu" title="Upload Photos">
										<IMG height="16" alt="Upload Photos to Workspace" src="<%=CMS_IMAGE_PATH%>/upload.gif" width="16" hspace="5" align="absmiddle">Upload Photos to Workspace
									</A>
								</span>
							</div>
						</div>
						<!-- END PAGE BODY TOOLBAR -->
						<table border="0" cellpadding="0" cellspacing="0" width="<%=CMS_PAGE_WIDTH%>" id="noworkspace"> 
							<tr>
								<td width="50" valign="top"><br>&nbsp;</td>
								<td style="width:auto;" align="left">
								<p style="padding-top:10;">There are not currently any photographs in the workspace which need to have 
								viewable photographs linked with their thumbnail representations. Please click 
								on the "Upload Photos to Workspace" link above to add photographs into this 
								workspace. You will need to choose whether the image is a Large ("Viewable") 
								photograph, or a thumbnail image.</p>
								<p>Once you have added photographs to this workspace, you will be able to view 
								them on this page and connect Large ("Viewable") photographs with their 
								thumbnail representations.</p>
								</td>
							</tr>
						</table>
						<table border="0" cellpadding="0" cellspacing="0" width="<%=CMS_PAGE_WIDTH%>" id="workspace"> 
							<tr>
								<td width="50" valign="top"><br>&nbsp;</td>
								<td style="width:auto;" align="center">
									<p class="admin-instruction"><% If IS_EDIT_PAGE Then echo(EDIT_INSTRUCTIONS) Else echo(REPORT_INSTRUCTIONS) %></p>
									<p style="color:red;"><%If StringNotEmptyOrNull(displayMessage) Then echo("<br>" & displayMessage)%></p>
									<table class="table-container" width="<%=CMS_PAGE_WIDTH%>">
										<tr><td style="border-bottom:1px solid #666666;color:#666;font-weight:bold;" id="largeImgCount"></td></tr>
										<tr><td>
											<div style="overflow:auto;height:225;width:<%=CMS_PAGE_WIDTH%>;SCROLLBAR-FACE-COLOR: #e7e7e7;SCROLLBAR-HIGHLIGHT-COLOR:#cccccc; SCROLLBAR-SHADOW-COLOR: #ccc; SCROLLBAR-3DLIGHT-COLOR: #FFFFFF; SCROLLBAR-ARROW-COLOR: #666666;SCROLLBAR-TRACK-COLOR: #FFFFFF; SCROLLBAR-DARKSHADOW-COLOR: #FFFFFF; SCROLLBAR-BASE-COLOR: #FFFFFF;">
											<% 
											dim hasViewableImages
											dim imgcounter
											hasViewableImages = false
											imgcounter = 0
											fsizecounter = 0
											Dim fs, f, f1, fc, imgName
											Set fs = CreateObject("Scripting.FileSystemObject") 
											
											'uncomment for localhost development
											'Set f = fs.GetFolder(Server.MapPath(strFsoFolder))  

											'die(strFsoFolder)

											'uncomment for production server
											Set f = fs.GetFolder(strFsoFolder)
											
											Set fc = f.files 
											Dim strStyle, bcolor,bgcolor
											Dim c,strType,ew,eh
											ew=0:eh=0
											For Each f1 in fc   

											    If lcase(f1.name) <> "thumbs.db" then
											    										
												    Call gfxSpex(f1.Path, ew, eh, c, strType)
												    imgName = LCase(left(f1.name,len(f1.name)-4))
												        												
												    If Contains(largePhotoGallery,imgName & ".jpg") = False Then

													    strStyle = strStyle & vbcrlf & "#Delete_" & imgname & "{visibility:hidden;}"									
													    Response.Write(vbcrlf & vbcrlf)
													    If Eval(CInt(ew) > CInt(MAX_PX_GALLERY_IMAGE_WIDTH)) or Eval(CInt(eh) > CInt(MAX_PX_GALLERY_IMAGE_HEIGHT)) then 
														    bcolor = "#ff0000"
														    bgcolor = "#ff0000"
													    else
														    bcolor = "#666666"
														    bgcolor = "#ffffff"
													    end If
													    if lcase(imgName) <> "thumb" then
														    hasViewableImages = true
														    imgcounter = imgcounter + 1
														    ' get the dimensions of the file. We pass arguments by reference to the gfxSpex function
														    fsizecounter = fsizecounter + f1.size
    													    
    													    
														    'Add the styles to a stylesheet and set to the className property
														    Response.write("<span id=""container_largeimg_" & lcase(imgName) & """ onmouseover=""javascript:this.style.backgroundColor='#F0E78C';this.style.borderColor='#F0E78C';toggleImageGallery('Delete_" & imgname & "');"" ")
														    response.write("onmouseout=""this.style.backgroundColor='" & bgcolor & "';this.style.borderColor='" & bcolor & "';toggleImageGallery('Delete_" & imgname & "');"" ")
														    response.write("style=""cursor:move;float:left;padding:1 8 1 8;border:1px solid " & bcolor & ";margin:2;background-color:" & bgcolor & ";"">") & vbCrLf
    	
														    'adding the delete capability
														    response.write("<div id=""Delete_" & imgName & """ align=""right"" width=""100%"" style=""color:#f00;margin-right:-6;"" ")
														    response.write("onclick=""location.href='photo-workspace.asp" & QSTRING & "&delete=" & f1.name & "&dtype=large';"">delete <img vspace=""0"" align=""top"" src=""" & CMS_IMAGE_PATH & "/ptclose.gif""></div>") & vbCrLf
    	
														    'Rendering the image
														    response.write("<img id=""largeimg_" & lcase(imgName) & """ ")
    														
														    'Open popup to preview image
														    response.write("onclick=""popupPhoto('common/popupPhoto.asp?Path=" & PHOTO_IMAGE_PATH & "/" & imgname & ".jpg&h=" & eh & "&w=" & ew & "','popup','','" & ew & "','" & eh & "','true');"" ")
														    response.write("src=""" & PHOTO_IMAGE_PATH & "/" & imgName & ".jpg"" width=""150"" height=""150"">") & vbCrLf
    														
														    'Display the dimensions
														    response.write("<div align=""center"" style=""font:10 'century gothic';color:#666;"">" & ew & " x " & eh)
														    response.Write(" <br><input type=""radio"" onclick=""radClick(this,'" & lcase(imgName) & "');"" name=""radLarge" & lcase(imgName) &  """ value=""rad_largeimg_" & lcase(imgName) & """>SELECT PHOTO")
														    response.write("</div></span>")
													    end if
    												End If
	                                            End If
											Next 
											%>
											</div>
											<script>
												var domLargeCounter = findDOM('largeImgCount');
												domLargeCounter.innerHTML='Large Images (<%=imgcounter%> images totaling <%=SizeFormat(fsizecounter)%>)';
											</script>												
										</td></tr>
									</table>
									<br>
									<select multiple name="selCnxn" id="selCnxn" size="1" style="display:none;"></select>
									<table border="0" cellpadding="0" cellspacing="0" width="550">
										<tr>
											<td width="350" align="right" valign="top">
                                                <!-- <table id="Table1" height="150" width="350" border="0" align="center" cellpadding="1" cellspacing="0" style="background:url(<%=CMS_IMAGE_PATH%>/ddbackdrop.gif) no-repeat top center;border:1 dotted #666;background-color:#ccc;" ondrop="drop()" ondragover="overDrag()" ondragenter="enterDrag()"> -->
												<table id="BackDrop" height="150" width="350" border="0" align="center" cellpadding="1" cellspacing="0" style="background:url(<%=CMS_IMAGE_PATH%>/ddbackdrop.gif) no-repeat top center;border:1px dotted #666666;background-color:#cccccc;">
													<tr>
														<td align="center" valign="bottom" height="125" width="175"><img id="largePreview" src="<%=CMS_IMAGE_PATH%>/LargeDrop.gif" height="100" width="100"></td>
														<td align="center"  valign="middle" height="125" width="175"><img id="thumbPreview" src="<%=CMS_IMAGE_PATH%>/ThumbDrop.gif" height="50" width="50"></td>
													</tr>
													<tr>
														<td id="LgName" style="font-weight:bold;line-height:15px;color:#666;" align="center">&nbsp;</td>
														<td id="ThbName" style="font-weight:bold;line-height:15px;color:#666;" align="center">&nbsp;</td>
													</tr>
												</table>											
											</td>
											<td width="200" align="left" valign="top" style="padding:5 0 0 5;">
												<div style="overflow:auto;width:175;height:150;SCROLLBAR-FACE-COLOR: #e7e7e7;SCROLLBAR-HIGHLIGHT-COLOR:#cccccc; SCROLLBAR-SHADOW-COLOR: #ccc; SCROLLBAR-3DLIGHT-COLOR: #FFFFFF; SCROLLBAR-ARROW-COLOR: #666666;SCROLLBAR-TRACK-COLOR: #FFFFFF; SCROLLBAR-DARKSHADOW-COLOR: #FFFFFF; SCROLLBAR-BASE-COLOR: #FFFFFF;">
													<a href="javascript:;" onclick="connectPhotos();viewConnectedPhotos();return false;" title="Connect the photos in the order in which you would like them to display."><img id="imgConnectPhotos" src="<%=CMS_IMAGE_PATH%>/butConnectPhotos.gif" style="border:1px solid #666666;" onMouseOver="this.style.bordercolor='#ffffff';" onMouseOut="this.style.bordercolor='#666666';" vspace="2"></a>																																							
													<a href="javascript:;" onclick="toggleMenuImage('imgViewConnectedPhotos','butViewConnectedPhotos.gif','butHideConnectedPhotos.gif');viewConnectedPhotos();toggle('tdConnectedPhotos');return false;" title="View/Hide all connected photos (Will toggle window on right-side of screen)"><img id="imgViewConnectedPhotos" src="<%=CMS_IMAGE_PATH%>/butViewConnectedPhotos.gif" style="border:1px solid #666666;" onMouseOver="this.style.bordercolor='#ffffff';" onMouseOut="this.style.bordercolor='#666666';" vspace="2"></a>
													<a href="javascript:;" onclick="resetConnections();return false;" title="Start over in the workspace"><img id="imgResetConnectedPhotos" src="<%=CMS_IMAGE_PATH%>/butResetConnectedPhotos.gif"style="border:1px solid #666666;" onMouseOver="this.style.bordercolor='#ffffff';" onMouseOut="this.style.bordercolor='#666666';" vspace="2"></a>
													<a href="#" onclick="javascript:submitForm();" title="Submit connected photos to save into database"><img id="imgSubmitConnectedPhotos" src="<%=CMS_IMAGE_PATH%>/butSubmitConnectedPhotos.gif" style="border:1px solid #666666;" onMouseOver="this.style.bordercolor='#ffffff';" onMouseOut="this.style.bordercolor='#666666';" vspace="2"></a>
												</div>
											</td>
										</tr>
									</table>
									<br>
									<table class="table-container" width="<%=CMS_PAGE_WIDTH%>">
										<tr><td style="color:#666;font-weight:bold;border-bottom:1px solid #666666;" id="thumbImgCount"></td></tr>
										<tr><td>
											<div style="overflow:auto;height:150;width:<%=CMS_PAGE_WIDTH%>;SCROLLBAR-FACE-COLOR:#e7e7e7;SCROLLBAR-HIGHLIGHT-COLOR:#cccccc; SCROLLBAR-SHADOW-COLOR: #ccc; SCROLLBAR-3DLIGHT-COLOR: #FFFFFF; SCROLLBAR-ARROW-COLOR: #666666;SCROLLBAR-TRACK-COLOR: #FFFFFF; SCROLLBAR-DARKSHADOW-COLOR: #FFFFFF; SCROLLBAR-BASE-COLOR: #FFFFFF;">
											<%	
											dim hasThumbImages
											imgcounter = 0
											hasThumbImages = false
											Set fs = CreateObject("Scripting.FileSystemObject")  
											
											'uncomment for localhost development
											'strFolder = Server.MapPath(strFsoFolder & "\thumbs")
											
											'uncomment for production server
											strFolder = strFsoFolder & "\thumbs"
											
											Set f = fs.GetFolder(strFolder)
											Set fc = f.files 
											Dim tfsizecounter
											ew=0:eh=0
											For Each f1 in fc   

                                                If lcase(f1.name) <> "thumbs.db" Then
                                                
												    Call gfxSpex(f1.Path, ew, eh, c, strType)
												    imgName = left(f1.name,len(f1.name)-4)

												    If Contains(thumbPhotoGallery,imgName & ".jpg") = False Then

													    strStyle = strStyle & vbcrlf & "#Delete_" & imgname & " {visibility:hidden;}"									
													    Response.Write(vbcrlf & vbcrlf)
													    If Cint(ew) > CInt(MAX_PX_THUMBNAIL_WIDTH) or CInt(eh) > CInt(MAX_PX_THUMBNAIL_HEIGHT) then 
														    bcolor = "#f00"
														    bgcolor = "#f00"
													    else
														    bcolor = "#666"
														    bgcolor = "#fff"
													    end If
    	
													    If lcase(imgName) <> "thumb" then
														    hasThumbImages = true								
														    imgcounter = imgcounter + 1
														    tfsizecounter = tfsizecounter + f1.size																								

														    'define in stylesheet and set className property
														    Response.write("<span id=""container_thumb_" & lcase(imgname) & """ onmouseover=""this.style.backgroundColor='#F0E78C';this.style.borderColor='#F0E78C';toggleImageGallery('Delete_" & imgname & "');"" ")
														    response.write("onmouseout=""this.style.backgroundColor='" & bgcolor & "';this.style.borderColor='" & bcolor & "';toggleImageGallery('Delete_" & imgname & "');"" ")
														    response.write("style=""cursor:move;float:left;padding:1 8 1 8;border:1px solid " & bcolor & ";margin:2;background-color:" & bgcolor & ";"">")
    		
														    'adding the delete capability
														    response.write("<div id=""Delete_" & imgName & """ align=""right"" width=""100%"" style=""color:#f00;margin-right:-6;"" ")
														    response.write("onclick=""location.href='photo-workspace.asp" & QSTRING & "&delete=" & f1.name & "&dtype=thumb';"">delete <img vspace=""0"" align=""top"" src=""" & CMS_IMAGE_PATH & "/ptclose.gif""></div>")
    		
														    'render the thumbnail
														    response.write("<img id=""thumb_" & lcase(imgName) & """ src=""" & PHOTO_IMAGE_PATH & "/thumbs/" & imgName & ".jpg"" width=""50"" height=""50"" >")
    														    		
														    'display the image dimensions
														    response.write("<div align=""center"" style=""font:10 'century gothic';color:#666;"">" & ew & " x " & eh)
                                                            response.Write(" <br><input type=""radio"" onclick=""radClick(this,'" & lcase(imgName) & "');"" name=""thumb_" & lcase(imgName) & """ value=""rad_thumb_" & imgName & """>SELECT PHOTO")
														    response.write("</div></span>")
    													
													    End If
												    End If
                                                End If
											Next 
											Response.Write(vbCrLf & vbCrLf & "<style>" & strStyle & vbCrLf & "</style>" & vbCrLf)
											%>
											</div>	
											<script>
												var domThumbCounter = findDOM('thumbImgCount');
												domThumbCounter.innerHTML='Thumbnail Images (<%=imgcounter%> images totaling <%=SizeFormat(tfsizecounter)%>)';
											</script>											
										</td></tr>
									</table>									
								</td>
								<td width="300px" valign="top" id="tdConnectedPhotos" style="width:310px;display:none;padding-top:9;" bgcolor="#dddddd">
									<div style="align:left;border-bottom:1px solid #666666;width:300px;padding-right:10px;color:#666;font-weight:bold;">Connected Photos</div>
									<div style="align:left;border-bottom:1px solid #666666;width:300px;padding:0 10 5 3;color:#666;">
										<span style="text-decoration:italics;font-weight:bold;color:#666;">Instructions:</span>
										The following photographs have been connected within the current workspace, but have not yet been submitted to 
										the database. Don't forget to submit all connected photographs the database, before leaving this page; otherwise, 
										your photograph relationships will be lost.<br><br>
										<span style="text-decoration:italics;font-weight:bold;color:#666;">Note:</span>To disassociate a thumbnail from the viewable image, please click on the image in the list below.
									</div>
									<div id="connectedPhotoList" style="overflow:auto;width:300px;height:450;padding:3 0 3 3;SCROLLBAR-FACE-COLOR:#e7e7e7;SCROLLBAR-HIGHLIGHT-COLOR:#cccccc; SCROLLBAR-SHADOW-COLOR: #ccc; SCROLLBAR-3DLIGHT-COLOR: #FFFFFF; SCROLLBAR-ARROW-COLOR: #666666;SCROLLBAR-TRACK-COLOR: #FFFFFF; SCROLLBAR-DARKSHADOW-COLOR: #FFFFFF; SCROLLBAR-BASE-COLOR: #FFFFFF;"></div>
								</td>
							</tr>
						</table>																			
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<!--#include virtual="/cms/common/__footer.asp"-->
</form>

<%
Erase largePhotoGallery
Erase thumbPhotoGallery

echo(vbCrLf & "<style><!--" & vbCrLf)
If (hasViewableImages = false And hasThumbImages = false) then
	'If there are no photos to display, then just show a blank page
	echo("#workspace {display:none;}" & vbCrLf)
	echo("#noworkspace {display:block;}" & vbCrLf)
Else
	echo("#workspace {display:block;}" & vbCrLf)
	echo("#noworkspace {display:none;}" & vbCrLf)
End If
echo(vbCrLf & "//--></style>" & vbCrLf)

%>

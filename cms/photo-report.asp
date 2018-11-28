<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/cms/common/__fsoConfig.asp" -->
<!-- #include virtual="/cms/common/__fsoCommon.asp" -->
<!-- #include virtual="/objects/cPhotos.asp" -->
<%
Dim PAGE_IMAGE, EDIT_PAGE, REPORT_PAGE, PAGE_TITLE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_desfield.gif"
PAGE_TITLE = "Photograph Listing"
'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim intPhotoId

Dim strPhotoType
Dim QSTRING 
Dim galleryId
Dim categoryId
Dim PHOTO_IMAGE_PATH
Dim PHOTO_IMAGE_MAPPED_PATH
Dim path,tpath
If StringNotEmptyOrNull(GetQryString("gid")) Then
	strPhotoType = "GALLERY"
	QSTRING = "?gid=" & GetQryString("gid") & "&gln=" & GetQryString("gln") & "&gn=" & GetQryString("gn")
	galleryId = GetQryString("gid")
	PHOTO_IMAGE_MAPPED_PATH = "../secure/proofs/" & LCase(PHOTOGRAPHER_FNAME) & "/" & GetQryString("gln")
	PHOTO_IMAGE_PATH = ROOT_PATH & PROOF_PATH & "/" & Replace(GetQryString("gln")," ","")
	
	'uncomment for production server
	path = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & GetAppVariable("PROOF_PATH") & "\" & GetQryString("gln") & "\" & GetQryString("large"),"/","\")
	tpath = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & GetAppVariable("PROOF_PATH") & "\" & GetQryString("gln") & "\thumbs\" & GetQryString("thumb"),"/","\")
	
	'uncomment for development on localhost
    'path = GetFilePath(PHOTO_IMAGE_PATH & "/" & GetQryString("large"))	
    'tpath = GetFilePath(PHOTO_IMAGE_PATH & "/thumbs/" & GetQryString("thumb"))	
Else
	strPhotoType = "CATEGORY"
	QSTRING = "?cid=" & GetQryString("cid") & "&cn=" & GetQryString("cn")
	categoryId = GetQryString("cid")
	PHOTO_IMAGE_MAPPED_PATH = "../secure/portfolio/" & LCase(PHOTOGRAPHER_FNAME) & "/" & GetQryString("cn")	
	PHOTO_IMAGE_PATH = ROOT_PATH & GALLERY_PATH & "/" & GetQryString("cn")
	
	'uncomment for production server
	path = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & GetAppVariable("GALLERY_PATH") & "\" & GetQryString("cn") & "\" & GetQryString("large"),"/","\")
	tpath = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & GetAppVariable("GALLERY_PATH") & "\" & GetQryString("cn") & "\thumbs\" & GetQryString("thumb"),"/","\")		

	'uncomment for development on localhost
    'path = GetFilePath(PHOTO_IMAGE_PATH & "/" & GetQryString("large"))	
    'tpath = GetFilePath(PHOTO_IMAGE_PATH & "/thumbs/" & GetQryString("thumb"))	
End If

    'echobr(replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & GetAppVariable("GALLERY_PATH") & "\" & GetQryString("cn") & "\" & GetQryString("large"),"/","\"))
    'echobr(GetFilePath(GALLERY_PATH & "/" & GetQryString("large")))
    'echobr("--")
    'echobr("PHOTO_IMAGE_MAPPED_PATH=" & PHOTO_IMAGE_MAPPED_PATH)
    'echobr("ROOT_PATH=" & ROOT_PATH)
    'echobr("GALLERY_PATH=" & GALLERY_PATH)
    'echobr("PHOTO_IMAGE_PATH=root_path & gallery_path which = " & PHOTO_IMAGE_PATH)
    'echobr(server.MapPath(gallery_path & "/" & getqrystring("cn")))
        
    'echobr("path= photo_image_path / and -> " & path)
    'echobr("tpath= photo_image_path / thumbs / and -> " & tpath)
    'echobr(path)
    'echobr(tpath)
    'die("end")
    
    
EDIT_PAGE = "photo-workspace.asp" & QSTRING 
REPORT_PAGE = "photo-report.asp" & QSTRING
IS_REPORT_PAGE = True

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function UpdatePhoto(pType,id,PhotoCaption,ImageOrder,intActiveFlag)
	Dim oPhoto
	Set oPhoto = New cPhotos
	With oPhoto	
		.ID = id
		.Caption = PhotoCaption
		.ImageOrder = ImageOrder
		.ActiveFlag = intActiveFlag
		If pType = "GALLERY" Then
			.UpdateGalleryPhotoInfo()
		ElseIf pType = "CATEGORY" Then
			.UpdateCategoryPhotoInfo()
		End If
		UpdatePhoto = Not .IsError
	End With
	Set oPhoto = Nothing
End Function

Function DeletePhoto(pType,id,largeFile,thumbFile)
	
	If StringNotEmptyOrNull(largeFile) Then
		Call DeleteFSOFile(largeFile)
	End If
	
	If StringNotEmptyOrNull(thumbFile) Then
		Call DeleteFSOFile(thumbFile)
	End If
	
	Dim oPhoto
	Set oPhoto = New cPhotos
	With oPhoto
		.ID = id
		If pType = "GALLERY" Then
			.DeleteGalleryPhotos()
		ElseIf pType = "CATEGORY" Then
			.DeleteCategoryPhotos()
		End If
		DeletePhoto = Not .IsError
	End With
	
End Function

Function DeReferencePhoto(pType,id)
	DeReferencePhoto = DeletePhoto(pType,id,"","")
End Function


'----------------------------------------------------------------------------------------
' PAGE RENDNER LOGIC
'----------------------------------------------------------------------------------------
intPhotoId = GetQryString("id")
Dim counter,frm,strCaption,intImageOrder,intActiveFlag,arrFrm
If intPhotoId <> 0 Then
	Select Case GetQryString("action")
		Case "delete"
		    'echobr(path)
		    'die(tpath)
		    
			blnSuccess = DeletePhoto(strPhotoType,intPhotoId,path,tpath)
			If blnSuccess Then		
				displayMessage = "Both the large viewable photograph and thumbnail representation were completely deleted from the file system and database."
			Else
				displayMessage = "An error occurred while attempting to delete the information from the server and/or database."
			End If
		Case "unlink"
			blnSuccess = DeReferencePhoto(strPhotoType,intPhotoId)
			If blnSuccess Then		
				displayMessage = "Both the large viewable photograph and thumbnail representation were successfully disconnected from each other. You can visit the photo workspace to manage those photographs."
			Else
				displayMessage = "An error occurred while attempting to delete the information from the database."
			End If
		Case "update"
			'strCaption = GetFormPost("taCaption_" & intPhotoId)
			intImageOrder = GetFormPost("txtImgOrder_" & intPhotoId)
			intActiveFlag = GetSqlCheckboxValue(GetFormPost("ckActiveFlag_" & intPhotoId))
			blnSuccess = UpdatePhoto(strPhotoType,GetQryString("id"),null,intImageOrder,intActiveFlag)
			If blnSuccess Then		
				displayMessage = "Your changes to photo information were successfully saved."
			Else
				displayMessage = "An error occurred while attempting to update your photo information to the database."
			End If		
		Case "saveall"
			For counter = 1 to Request.Form.Count
				frm = Request.Form.Key(counter)
				If Instr(frm,"_") > 0 Then
					arrFrm = Split(frm,"_")
					intPhotoId = Trim(arrFrm(1))
					strCaption = GetFormPost("taCaption_" & intPhotoId)
					intImageOrder = GetFormPost("txtImgOrder_" & intPhotoId)
					intActiveFlag = GetSqlCheckboxValue(GetFormPost("ckActiveFlag_" & intPhotoId))
					blnSuccess = UpdatePhoto(strPhotoType,intPhotoId,strCaption,intImageOrder,intActiveFlag)
					If blnSuccess = False Then
						Exit For
					End If
				End If
			Next
			If blnSuccess Then		
				displayMessage = "All changes to the current page were successfully saved."
			Else
				displayMessage = "An error occurred while attempting to update the batch of information to the database. (Error with Photo Id: " & intPhotoId & ")"
			End If
	End Select
End If
%>
<HTML>
<HEAD>
<TITLE><%=PAGE_TITLE%></TITLE>
<script src="/cms/lib/__dom.js" type="text/javascript"></script>
<script><!--
function checkForm(type,id) {
	var frm = document.forms['form1'];
	frm.action = '<%=REPORT_PAGE%>&id='+id+'&action='+type;
	frm.submit();
}
//--></script>
<!--#include virtual="/cms/styles.asp"-->
</HEAD>
<BODY>
<form name="form1" id="form1" method="post" action="<%=EDIT_PAGE%>">
<input type="hidden" name="hidMode" id="hidMode" value="<%=strMode%>">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" style="border-bottom-style:solid;border-width:1px;border-color:#666666;">
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
								<span style="float:right;padding-right:10px;">
									<A href="#" onClick="return checkForm('saveall',-1);" class="menu" title="Save and Close">
										<IMG height="16" alt="Save All and Reload" src="<%=CMS_IMAGE_PATH%>/savereload.gif" width="16" hspace="5" align="absmiddle">Save All and Reload
									</A>
								</span>
								<span style="float:right;padding-right:10px;">
									<a href="<%=EDIT_PAGE%>" class="menu" title="Add/Upload Photos">
										<img src="<%=CMS_IMAGE_PATH%>/addimage.gif" alt="Upload/Link Photos for this gallery" hspace="5" align="absmiddle">Add/Upload Photos
									</a>
								</span>
							</div>
						</div>
						<!-- END PAGE BODY TOOLBAR -->
						<table border="0" cellpadding="0" cellspacing="0" width="<%=CMS_PAGE_WIDTH%>">
							<tr>
								<td width="50" valign="top"><br><img src="<%=CMS_IMAGE_PATH %>/<%=PAGE_IMAGE%>"></td>
								<td style="width:auto;">
									<p class="admin-instruction"><% If IS_EDIT_PAGE Then echo(EDIT_INSTRUCTIONS) Else echo(REPORT_INSTRUCTIONS) %></p>
									<p style="color:red;"><%If StringNotEmptyOrNull(displayMessage) Then echo("<br>" & displayMessage)%></p>										
									<table class="report" width="100%" cellpadding="0" cellspacing="0">
										<tr>
											<td class="report-header">&nbsp;</td>
											<td class="report-header">Thumbnail</td>
											<td class="report-header">Viewable</td>
											<!--<td class="report-header">Caption</td>-->
											<td class="report-header">Order</td>
											<td class="report-header" align="center" width="100"><a href="#" id="hrefToggleActiveInd" name="hrefToggleActiveInd" onclick="ToggleAll(this);" class="report-header" style="text-decoration:underline;">Activate</a></td>
											<td class="report-header">&nbsp;</td>
										</tr>
										<%				
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
										%>
											<tr>
												<td>
													<a href="<%=REPORT_PAGE%>&id=<%=oPhoto.ID%>&action=delete&large=<%=oPhoto.LargeImage%>&thumb=<%=oPhoto.ThumbImage%>"><img src="<%=CMS_IMAGE_PATH%>/delitem.gif" vspace="10" alt="Delete"></a>&nbsp;
													<a href="<%=REPORT_PAGE%>&id=<%=oPhoto.ID%>&action=unlink"><img src="<%=CMS_IMAGE_PATH%>/unlink.gif" vspace="10" alt="Unlink Photographs" width="16" height="16"></a>
												</td>
												<td><br><img src="<%=PHOTO_IMAGE_PATH%>/thumbs/<%=oPhoto.ThumbImage%>" height="50" width="50"><br><%=oPhoto.ThumbImage%></td>	
<%												
														Dim strFile,objFile, ew,eh,depth,strType

														'Obtain physical path to the viewable graphic file
														'uncomment for production server
                                                        If strPhotoType = "GALLERY" Then
                                                            strFile = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & GetAppVariable("PROOF_PATH") & "\" & GetQryString("gln") & "\" & oPhoto.LargeImage,"/","\")
                                                        Else 'CATEGORY
                                                            strFile = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & GetAppVariable("GALLERY_PATH") & "\" & GetQryString("cn") & "\" & oPhoto.LargeImage,"/","\")
                                                        End If
														'strFile = path & oPhoto.LargeImage
														
														'uncomment for development on localhost
														'strFile = Server.MapPath(PHOTO_IMAGE_MAPPED_PATH & "/" &  oPhoto.LargeImage)
																												
														'Get object reference to the viewable file -- used to get spec information
														Set objFile = objFSO.GetFile(strFile)
														
														'Obtain height/width properties for current image file
														Call gfxSpex(objFile,ew,eh,depth,strType)
														
														'echobr(IsChecked(oPhoto.ActiveFlag))
														'die(oPhoto.ActiveFlag)
%>
												<td><br><img style="cursor:hand;" src="<%=PHOTO_IMAGE_PATH%>/<%=oPhoto.LargeImage%>" height="100" width="100" onclick="popupPhoto('common/popupPhoto.asp?Path=<%=PHOTO_IMAGE_PATH & "/" & oPhoto.LargeImage & "&h=" & eh & "&w=" & ew%>','popup','','<%=ew%>','<%=eh%>','true');"><br><%=oPhoto.LargeImage%></td>													
												<!--<td><textarea cols="40" rows="4" wrap="virtual" id="taCaption_<%=oPhoto.ID%>" name="taCaption_<%=oPhoto.ID%>"><%=oPhoto.Caption%></textarea></td>-->
												<td align="center"><input type="text" id="txtImgOrder_<%=oPhoto.ID%>" name="txtImgOrder_<%=oPhoto.ID%>" size="5" value="<%=oPhoto.ImageOrder%>"></td>
												<td align="center"><input type="checkbox" name="ckActiveFlag_<%=oPhoto.ID%>" <%=IsChecked(oPhoto.ActiveFlag)%>></td>
												<td><A href="#" onClick="return checkForm('update',<%=oPhoto.ID%>);" class="menu" title="Save Photo Information"><IMG height="16" alt="Update Photo Info" src="<%=CMS_IMAGE_PATH%>/savesingleitem.gif" width="16" hspace="5" align="absmiddle"></A></td>
											</tr>									
										<%
											Set oPhoto = Nothing
										Next
										Set collPhotos = Nothing
										%>
									</table>																	
	<!--#include virtual="/cms/common/__end_bodywrap.asp"-->
</table>
<!--#include virtual="/cms/common/__footer.asp"-->
</BODY>
</form>
</HTML>

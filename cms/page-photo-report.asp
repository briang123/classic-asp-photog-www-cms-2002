<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/cms/common/__fsoConfig.asp" -->
<!-- #include virtual="/cms/common/__fsoCommon.asp" -->
<!-- #include virtual="/objects/cPhotos.asp" -->
<!-- #include virtual="/objects/cPageInfo.asp" -->
<%
Dim PAGE_IMAGE, EDIT_PAGE, REPORT_PAGE, PAGE_TITLE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_desfield.gif"
PAGE_TITLE = "Photograph Listing"

Dim filterQString
If StringNotEmptyOrNull(GetQryString("fid")) Then
	If GetQryString("fid") = 0 Then
		filterQString = "'&fid=0'"
	Else											
		filterQString = "'&fid=" & GetQryString("fid") & "'"
	End If
Else
	filterQString = "''"
End If

Public Function GetPageInfo(ctlName,val)

	Dim tempStr
	
	tempStr = "<select name=""" & ctlName & """ id=""" & ctlName & """>"
	tempStr = tempStr & "<option value=""0"""
	If StringEmptyOrNull(val) Then
		tempStr = tempStr & " selected"
	End If
	tempStr = tempStr & ">---SELECT ONE---</option>"

	Dim collPageInfo, oPageInfo
	Set oPageInfo = New cPageInfo
	Set collPageInfo = New cPageInfo
	Call collPageInfo.GetPageInfo()
	For Each oPageInfo In collPageInfo.PageInfo.Items
		tempStr = tempStr & "<option value=""" & oPageInfo.ID & """"
		If CInt(oPageInfo.ID) = CInt(val) Then
			tempStr = tempStr & " selected>"
		Else
			tempStr = tempStr & ">"
		End If
		tempStr = tempStr & QuoteCleanup(oPageInfo.WebPage) & "</option>"
	Next
	Set oPageInfo = Nothing
	Set collPageInfo = Nothing
	tempStr = tempStr & "</select>"
	
	GetPageInfo = tempStr
	
End Function

'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim intPhotoId

Dim PHOTO_IMAGE_PATH
Dim PHOTO_IMAGE_MAPPED_PATH
PHOTO_IMAGE_MAPPED_PATH = ".." & IMAGE_PATH & "/"
PHOTO_IMAGE_PATH = IMAGE_PATH

'echobr(replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & GetAppVariable("IMAGE_PATH"),"/","\"))
'echobr(PHOTO_IMAGE_MAPPED_PATH)
'echobr(PHOTO_IMAGE_PATH)

EDIT_PAGE = "common/__siteUploadForm.asp?fid=" & GetQryString("fid")
REPORT_PAGE = "page-photo-report.asp"
IS_REPORT_PAGE = True

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function UpdatePhoto(id,PageId,ImageOrder,intActiveFlag)
	Dim oPhoto
	Set oPhoto = New cPhotos
	With oPhoto	
		.ID = id
		.WebPageId = PageId
		.ImageOrder = ImageOrder
		.ActiveFlag = intActiveFlag
		.UpdateSitePhotoInfo()
		UpdatePhoto = Not .IsError
	End With
	Set oPhoto = Nothing
End Function

If StringNotEmptyOrNull(Session("SiteDeleteMessage")) then
	displayMessage = Session("SiteDeleteMessage")
else
	displayMessage = ""
end if


Function DeletePhoto(id,largeFile)
		
	If StringNotEmptyOrNull(largeFile) Then
		Call DeleteFSOFile(largeFile)
	End If
	
	Dim oPhoto
	Set oPhoto = New cPhotos
	With oPhoto
		.ID = id
		.DeleteSitePhotos()
		blnSuccess = Not .IsError
	End With

	If blnSuccess Then		
		Session("SiteDeleteMessage") = "The site photograph was completely deleted the system."
	Else
		Session("SiteDeleteMessage") = "An error occurred while attempting to delete the information from the server and/or database."
	End If

	Response.Redirect(REPORT_PAGE & "?fid=" & GetQryString("fid"))
	
End Function

Private Function File(byVal pathName)
	Dim objFSO
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	File = objFSO.FileExists(pathName)
	Set objFSO = Nothing
End Function


'Call AppendDebugInfo("page name","page-photo-report.asp")

'----------------------------------------------------------------------------------------
' PAGE RENDNER LOGIC
'----------------------------------------------------------------------------------------
intPhotoId = GetQryString("id")
Dim counter,frm,intImageOrder,intActiveFlag,arrFrm,intPageId,exists
If intPhotoId <> 0 Then
	Select Case GetQryString("action")
		Case "delete"	    	

'Call AppendDebugInfo("action type","delete")
		    'uncomment for localhost development
			'exists = File(Server.MapPath(PHOTO_IMAGE_PATH & "/" & GetQryString("large")))
			
			'uncomment for production server
	    	Dim path
	    	path = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & GetAppVariable("IMAGE_PATH") & "\" & GetQryString("large"),"/","\")

''Call AppendDebugInfo("path",path)


			exists = File(path)
			
'Call AppendDebugInfo("file exists",exists)



			if exists then
			
			    'uncomment for localhost development
				'blnSuccess = DeletePhoto(intPhotoId,Server.MapPath(PHOTO_IMAGE_PATH & "/" & GetQryString("large")))
				
'Call AppendDebugInfo("attempting to delete photo",intPhotoId & "::" & path)


				'uncomment for production server
				blnSuccess = DeletePhoto(intPhotoId,path)

'Call AppendDebugInfo("delete successful",blnSuccess)

				
				'If blnSuccess Then		
				'	displayMessage = "The site photograph was completely deleted the system."
				'Else
				'	displayMessage = "An error occurred while attempting to delete the information from the server and/or database."
				'End If
				
				
			end if
		Case "update"

'Call AppendDebugInfo("action type","update")

			intImageOrder = GetFormPost("txtImgOrder_" & intPhotoId)
			intActiveFlag = GetSqlCheckboxValue(GetFormPost("ckActiveFlag_" & intPhotoId))
			intPageId = GetFormPost("selPageInfo_" & intPhotoId)
			blnSuccess = UpdatePhoto(intPhotoId,intPageId,intImageOrder,intActiveFlag)
'Call AppendDebugInfo("update successful",blnSuccess)
			If blnSuccess Then		
				displayMessage = "Your changes to photo information were successfully saved."
			Else
				displayMessage = "An error occurred while attempting to update your photo information to the database."
			End If		
		Case "saveall"

'Call AppendDebugInfo("action type","save all")

			For counter = 1 to Request.Form.Count
				frm = Request.Form.Key(counter)
				If Instr(frm,"_") > 0 Then
					arrFrm = Split(frm,"_")
					intPhotoId = Trim(arrFrm(1))
					intImageOrder = GetFormPost("txtImgOrder_" & intPhotoId)
					intActiveFlag = GetSqlCheckboxValue(GetFormPost("ckActiveFlag_" & intPhotoId))
					intPageId = GetFormPost("selPageInfo_" & intPhotoId)
					blnSuccess = UpdatePhoto(intPhotoId,intPageId,intImageOrder,intActiveFlag)


					If blnSuccess = False Then

'Call AppendDebugInfo("save all failed","YES")
'Call AppendDebugInfo("photo id",intPhotoId)
'Call AppendDebugInfo("page id",intPageId)

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
function checkForm(type,id,filter) {
	var pgId = findDOM('hidPageId');
	(filter==''||parseInt(filter)==NaN) ? pgId.value=0 : pgId.value=parseInt(filter);
	var frm = document.forms['form1'];
	frm.action = '<%=REPORT_PAGE%>?id='+id+'&action='+type+filter.replace('\'','');
	frm.submit();
}
//--></script>
<!--#include virtual="/cms/styles.asp"-->
</HEAD>
<BODY>
<form name="form1" id="form1" method="post" action="<%=EDIT_PAGE%>">
<input type="hidden" name="hidMode" id="hidMode" value="<%=strMode%>">
<input type="hidden" name="hidPageId" id="hidPageId" value="0">
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
									<A href="#" onClick="return checkForm('saveall',-1,<%=filterQString%>);" class="menu" title="Save and Close">
										<IMG height="16" alt="Save All and Reload" src="<%=CMS_IMAGE_PATH%>/savereload.gif" width="16" hspace="5" align="absmiddle">Save All and Reload
									</A>
								</span>
								<span style="float:right;padding-right:10px;">
									<A href="#" onClick="javascript:popup('<%=EDIT_PAGE%>','Upload',500,500);return false;" class="menu" title="Upload Photos">
										<IMG height="16" alt="Upload Site Photos" src="<%=CMS_IMAGE_PATH%>/upload.gif" width="16" hspace="5" align="absmiddle">Upload Site Photos
									</A>
									
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
											<td class="report-header" align="center">Site Photo</td>
											<td class="report-header" align="center">Order</td>
											<td class="report-header" align="center">Web Page</td>
											<td class="report-header" align="center" width="100"><a href="#" id="hrefToggleActiveInd" name="hrefToggleActiveInd" onclick="ToggleAll(this);" class="report-header" style="text-decoration:underline;">Activate</a></td>
											<td class="report-header">&nbsp;</td>
										</tr>
										<%
'										Dim filterQString				
										Set oPhoto = New cPhotos
										Set collPhotos = New cPhotos
										
										If StringNotEmptyOrNull(GetQryString("fid")) Then
											If GetQryString("fid") = 0 Then
'												filterQString = "'&fid=0'"
												Call collPhotos.GetSitePhotos()
											Else											
'												filterQString = "'&fid=" & GetQryString("fid") & "'"
												collPhotos.WebPageId = GetQryString("fid")
												Call collPhotos.GetSitePhotosByPage()
											End If
										Else
											filterQString = "''"
											Call collPhotos.GetSitePhotos()										
										End If
										
										For Each oPhoto In collPhotos.Photos.Items
										%>
											<tr<% If oPhoto.ActiveFlag = False Then echo(" bgcolor=""#aaaaaa""")%>>
												<td>
													<a href="<%=REPORT_PAGE%>?id=<%=oPhoto.ID%>&action=delete&large=<%=oPhoto.LargeImage%>&fid=<%=GetQryString("fid")%>"><img src="<%=CMS_IMAGE_PATH%>/delitem.gif" vspace="10" alt="Delete"></a>
												</td>
												<%
														Dim strFile,objFile, ew,eh,depth,strType

														'Obtain physical path to the viewable graphic file
														'uncomment for production server
														strFile = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & GetAppVariable("IMAGE_PATH"),"/","\") & "\" & oPhoto.LargeImage
														

'Call AppendDebugInfo("photo",strFile)


														'uncomment for localhost development
														'strFile = Server.MapPath(IMAGE_PATH & "/" & oPhoto.LargeImage)
'echobr(strFile)

														'Get object reference to the viewable file -- used to get spec information
														Set objFile = objFSO.GetFile(strFile)
														
														'Obtain height/width properties for current image file
														Call gfxSpex(objFile,ew,eh,depth,strType)
                                                %>
												<td><br><img style="border:1px solid #666666;cursor:hand;" src="<%=PHOTO_IMAGE_PATH%>/<%=oPhoto.LargeImage%>" height="100" width="100" onclick="popupPhoto('common/popupPhoto.asp?Path=<%=PHOTO_IMAGE_PATH & "/" & oPhoto.LargeImage & "&h=" & eh & "&w=" & ew%>','popup','','<%=ew%>','<%=eh%>','true');"><br><%=oPhoto.LargeImage%></td>													
												<td align="center"><input type="text" id="txtImgOrder_<%=oPhoto.ID%>" name="txtImgOrder_<%=oPhoto.ID%>" size="5" value="<%=oPhoto.ImageOrder%>"></td>
												<td align="center"><%=GetPageInfo("selPageInfo_" & oPhoto.ID,oPhoto.WebPageId)%></td>
												<td align="center"><input type="checkbox" name="ckActiveFlag_<%=oPhoto.ID%>" <%=IsChecked(oPhoto.ActiveFlag)%>></td>
												<td><A href="#" onClick="return checkForm('update',<%=oPhoto.ID%>,'&fid=<%=GetQryString("fid")%>');" class="menu" title="Save Photo Information"><IMG height="16" alt="Update Photo Info" src="<%=CMS_IMAGE_PATH%>/savesingleitem.gif" width="16" hspace="5" align="absmiddle"></A></td>
											</tr>									
										<%
											'SendDebugInfo

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

<%
Session("SiteDeleteMessage") = ""
%>
<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<%
Server.ScriptTimeout = SCRIPT_TIMEOUT_IN_MINUTES * 60

Dim GalleryLastName
GalleryName = ""
CategoryName = ""
GalleryLastName = LCase(GetQryString("gln"))
CategoryName = LCase(GetQryString("cn"))

Dim LargeImagePath
Dim ThumbImagePath
If StringNotEmptyOrNull(GalleryLastName) Then 
	LargeImagePath = "\" & Replace(PROOF_PATH,"/","\") & "\" & GalleryLastName
ElseIf StringNotEmptyOrNull(CategoryName) Then 
	LargeImagePath = "\" & Replace(GALLERY_PATH,"/","\") & "\" & CategoryName
End If
ThumbImagePath = LargeImagePath & "\thumbs"
%>
<html>
<head>
<title>Photo Upload Center</title>
<script language="JavaScript1.2" type="text/javascript">
<!--
// Adds a new file attachment component dynamically on the screen
var nfiles=<%=FILE_UPLOAD_BATCH_COUNT%>;
function Expand() {
	var adh = '<table width="440" border="0" style="border-bottom:dotted 1 #666;">';
<% 
Dim fileUploads
For fileUploads = 1 to FILE_UPLOAD_BATCH_COUNT 
%>
	nfiles++;
	adh += '<tr><td width="100" style="font-size:12px;">Photo '+nfiles+'</td><td width="320">';
	adh += '<input type="file" size="40" name="file'+nfiles+'" style="font:8pt verdana,arial,sans-serif"></td></tr>';	
<% Next %>
	adh += '</table>';
	files.insertAdjacentHTML('BeforeEnd',adh);
	return false;
}
function Submit() {
	var sel=document.getElementById('selPhotoType');
	if (!sel) {
		var sel=document.formUpload.elements["selPhotoType"];
	}
	document.formUpload.action='__uploadProcessor.asp?p='+sel.options[sel.selectedIndex].value;
	document.formUpload.submit();
}
//-->
</script>
<!--#include virtual="/cms/styles.asp"-->
<!--#include virtual="/stylesA.asp"-->
</head>
<body style="background-Color:#C2BF98;">
<table border="0" cellpadding="0" cellspacing="0" bgcolor="#C2BF98">
	<tr>
		<td style="padding:5;">
			<table width="475" cellpadding="0" cellspacing="0" border="0" bgcolor="#ffffff">
				<tr>
					<td bgcolor="#666666"><img src="<%=CMS_IMAGE_PATH%>/imagegallery.gif"></td>
				</tr>
				<tr>
					<td align="center">
					<% select case request.QueryString("msg")							
							case "fail"
								response.write "The upload process failed.</br>"
								Response.Write("<script>window.opener.location.reload();</script>")								
							case "success"
								response.write "The upload process was successful.<br/>"
								Response.Write("<script>window.opener.location.reload();</script>")
						end select 
					%>
					</td>
				</tr>
				<form name="formUpload" id="formUpload" method="post" enctype="multipart/form-data" action="__uploadProcessor.asp">
				<tr>
					<td height="5"><img src="<%=CMS_IMAGE_PATH%>/1x1.gif" width="1" height="5" border="0"></td>
				</tr>
				<tr>
					<td align="left">
						<div style="border-bottom:1 solid <%=abelard_border_color%>;width:475px;padding-bottom:2px;" align="center">
							<table width="475" align="center" cellpadding="0" cellspacing="0" border="0">
								<tr>
									<td align="left" width="50%"><!--
										<A href="javascript:void(0);" onClick="return Expand();" title="Add Upload Field" style="text-decoration:none;color:#000;">
											<IMG alt="Add Upload Fields" src="<%=CMS_IMAGE_PATH%>/addimage.gif" hspace="5" align="absmiddle">Add Upload Fields
										</A>-->						
									</td>
									<td align="right" width="50%">
										<select name="selPhotoType">
											<option value="<%=LargeImagePath%>">Large Image</option>
											<option value="<%=ThumbImagePath%>">Thumbnail</option>
										</select>									
									</td>
								</tr>
							</table>
						</div>					
					</td>
				</tr>
				<tr>
					<td align="left" style="padding:2 3 2 3;font:11 ""century gothic""">Please make your selection above of the type of photograph you are wishing to upload. You will not be permitted to mix & match thumbnail images and larger images in the same upload process.</td>
				</tr>
				<tr>
					<td style="padding:10 0 0 10;">	
						<div id="files" style="overflow:auto;height:300px;width:475px;SCROLLBAR-FACE-COLOR: #e7e7e7;SCROLLBAR-HIGHLIGHT-COLOR:#cccccc; SCROLLBAR-SHADOW-COLOR: #ccc; SCROLLBAR-3DLIGHT-COLOR: #FFFFFF; SCROLLBAR-ARROW-COLOR: #666666;SCROLLBAR-TRACK-COLOR: #FFFFFF; SCROLLBAR-DARKSHADOW-COLOR: #FFFFFF; SCROLLBAR-BASE-COLOR: #FFFFFF;">
							<table width="440" border="0" style="border-bottom:dotted 1 #666;">
								<% For fileUploads = 1 to FILE_UPLOAD_BATCH_COUNT %>
								<tr>
									<td width="100" style="font-size:12px;">Photo <%=fileUploads%></td>
									<td width="320"><input type="file" size="40" name="file<%=fileUploads%>" style="font:8pt verdana,arial,sans-serif"></td>
								</tr>
								<% Next %>
							</table>
						</div>
					</td>
				</tr>
				<tr>
					<td align="center" style="padding:10 0 5 0;">
						<input type="hidden" name="hidUploadPath" id="hidUploadPath" value="<%echo(GetAppVariable("PHYSICAL_ROOT_PATH")) 'localhost development --> PHYSICAL_ROOT_PATH%>">
						<input type="hidden" name="hidGalleryLastName" id="hidGalleryLastName" value="<%=GetQryString("gln")%>">
						<input type="hidden" name="hidCategoryName" id="hidCategoryName" value="<%=GetQryString("cn")%>">
						<input type="submit" value="Upload Files" style="height: 22px;font:8pt verdana,arial,sans-serif" onClick="Submit();">
						<input type="button" value="Cancel" onClick="window.close();" style="height: 22px;font:8pt verdana,arial,sans-serif">
					</td>
					</form> 
				</tr>
			</table>			
		</td>
	</tr>
</table>
</body>
</html>

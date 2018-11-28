<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/cms/common/__fsoConfig.asp" -->
<!-- #include virtual="/cms/common/__fsoCommon.asp" -->
<!-- #include virtual="/objects/cGallery.asp" -->
<!-- #include virtual="/objects/cLogin.asp" -->

<%
Dim HasNoErrors
HasNoErrors = True

Dim PAGE_IMAGE, EDIT_PAGE, REPORT_PAGE, PAGE_TITLE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_desfield.gif"
PAGE_TITLE = "Proofing Gallery"
EDIT_PAGE = "gallery.asp"
REPORT_PAGE = "gallery-report.asp"
IS_REPORT_PAGE = False
Dim PHOTO_IMAGE_PATH
'PHOTO_IMAGE_PATH = CMS_ROOT_PATH & PROOF_PATH & "/"
PHOTO_IMAGE_PATH = ROOT_PATH & PROOF_PATH & "/"

Public Function GetUsersCombo(ctlName,val)
	
	Dim tempStr
	
	tempStr = "<select name=""" & ctlName & """ id=""" & ctlName & """>"
	tempStr = tempStr & "<option value=""0"""
	If StringEmptyOrNull(val) Then
		tempStr = tempStr & " selected"
	End If
	tempStr = tempStr & ">---SELECT ONE---</option>"

	Dim collLogin, oLogin
	Set oLogin = New cLogin
	Set collLogin = New cLogin
	Call collLogin.GetUsers()
	For Each oLogin In collLogin.Logins.Items
		tempStr = tempStr & "<option value=""" & oLogin.ID & """"
		If oLogin.ID = val Then
			tempStr = tempStr & " selected>"
		Else
			tempStr = tempStr & ">"
		End If
		tempStr = tempStr & QuoteCleanup(oLogin.FullName) & "</option>"
		Set oLogin = Nothing
	Next
	Set collLogin = Nothing
	tempStr = tempStr & "</select>"
	
	GetUsersCombo = tempStr
	
End Function

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function AddGallery(GalleryLastName,GalleryName,ExpirationDate,UserId,intActiveFlag,lngGalleryId)
	
	Set fso = CreateObject("Scripting.FileSystemObject")

	'uncomment for localhost development
	'If fso.FolderExists(Server.MapPath(Replace(PHOTO_IMAGE_PATH & LCase(GalleryLastName) & "/","/","\"))) Then
	
	'uncomment for production server
	If fso.FolderExists(Server.MapPath(Replace("/" & PROOF_PATH & "/" & LCase(GalleryLastName) & "/","/","\"))) Then
		AddGallery = False
	Else
	    	'uncomment for localhost development
		'objFSO.CreateFolder(GetFilePath(PHOTO_IMAGE_PATH & LCase(GalleryLastName)))
		'objFSO.CreateFolder(GetFilePath(PHOTO_IMAGE_PATH & LCase(GalleryLastName) & "\thumbs"))
		
		'uncomment for production server
		'objFSO.CreateFolder(GetFilePath("/" & PROOF_PATH & "/" & LCase(GalleryLastName)))
		'objFSO.CreateFolder(GetFilePath("/" & PROOF_PATH & "/" & LCase(GalleryLastName) & "\thumbs"))

   		Dim path,tpath
    		path = Server.MapPath(Replace("/" & PROOF_PATH & "/" & LCase(GalleryLastName) & "/","/","\"))
    		tpath = Server.MapPath(Replace("/" & PROOF_PATH & "/" & LCase(GalleryLastName) & "/thumbs","/","\"))

		if Not fso.FolderExists(path) Then
	    		fso.CreateFolder(path)
	    		fso.CreateFolder(tpath)	
	    		AddGallery = True
		else
	    		AddGallery = False
    		end If
		Set objFSO = Nothing

		Dim oGallery
		Set oGallery = New cGallery
		With oGallery
			.GalleryLastName = GalleryLastName
			.GalleryName = GalleryName
			.ExpirationDate = ExpirationDate
			.GalleryUserId = UserId
			.ActiveFlag = intActiveFlag
			.AddGallery()
			lngGalleryId = .ID
			AddGallery = Not .IsError
		End With
		Set oGallery = Nothing
	End If

End Function

Function UpdateGallery(id,OldLastName,GalleryLastName,GalleryName,ExpirationDate,UserId,intActiveFlag)

'die(StrComp(OldLastName,GalleryLastName,0) <> 0)

	'If StrComp(OldLastName,GalleryLastName,0) <> 0 Then


	Dim oGallery
	Set oGallery = New cGallery
	With oGallery
		.ID = id
		.GalleryLastName = GalleryLastName
		.GalleryName = GalleryName
		.ExpirationDate = ExpirationDate
		.GalleryUserId = UserId
		.ActiveFlag = intActiveFlag
		.UpdateGallery()
		UpdateGallery = Not .IsError
	End With
	Set oGallery = Nothing

	Set fso = CreateObject("Scripting.FileSystemObject")
	if not fso.FolderExists(Server.MapPath(Replace("/" & PROOF_PATH & "/" & GalleryLastName & "/","/","\"))) then

		If UpdateGallery Then
			If StrComp(OldLastName,GalleryLastName,0) <> 0 Then
				Set fso2 = CreateObject("Scripting.FileSystemObject")
				Set fldr2 = objFSO.GetFolder(Server.MapPath(Replace("/" & PROOF_PATH & "/" & OldLastName & "/","/","\")))
				fldr2.Name = GalleryLastName
				Set fldr2 = Nothing
				Set fso2 = Nothing	
			End If
		End If	

		'else
			'UpdateGallery = False

	End If

		Set fso = Nothing	

	'Else
	'	UpdateGallery = True
	'End If


End Function

'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim lngGalleryId
Dim strGalleryName
Dim strGalleryLastName
Dim dtExpirationDate
Dim intGalleryUser
Dim intActiveFlag

dtExpirationDate = FormatDate(Now()+DAYS_FROM_NOW_TO_EXPIRE_ACCT,"%m/%d/%Y")

'----------------------------------------------------------------------------------------
' PAGE RENDER LOGIC
'----------------------------------------------------------------------------------------
If GetFormPost("hidMode") = "save" Then
	strMode = "edit"	

	'strOldGalleryLastName = GetFormPost("hidOldGalleryLastName")
	strOldGalleryLastName = ShrinkText(LCase(Replace(GetFormPost("hidOldGalleryLastName"),"'","")))
	strGalleryLastName = ShrinkText(LCase(Replace(GetFormPost("txtGalleryLastName"),"'","")))

	if StringEmptyOrNull(Session("OldGalleryText")) then
		Session("OldGalleryText") = strOldGalleryLastName  & ""
	end if

	lngGalleryId = GetFormPost("hidGalleryId")
	strGalleryName = QuoteCleanup(GetFormPost("txtGalleryName"))
	dtExpirationDate = GetFormPost("txtExpirationDate")
	intGalleryUser = GetFormPost("txtUserId")
	intActiveFlag = GetSqlCheckboxValue(GetFormPost("ckActiveFlag"))
	
	If StringNotEmptyOrNull(lngGalleryId) Then
		blnSuccess = UpdateGallery(lngGalleryId,Session("OldGalleryText"),strGalleryLastName, strGalleryName, dtExpirationDate,intGalleryUser,intActiveFlag)		
	Else
		blnSuccess = AddGallery(strGalleryLastName, strGalleryName,dtExpirationDate,intGalleryUser,intActiveFlag,lngGalleryId)
	End If
		
	If blnSuccess Then
		Session("OldGalleryText") = ""
		PageRedirect(REPORT_PAGE)
	Else
		displayMessage = "An error occurred while trying to save information to the database or file system. It could be that you were trying to create a gallery which already exists."
	End If
Else
    strMode = "save"
	lngGalleryId = GetQryString("id")
	If StringNotEmptyOrNull(lngGalleryId) Then
		Set oGallery = New cGallery
		Set collGallery = New cGallery
		collGallery.ID = lngGalleryId
		Call collGallery.GetGalleryById()
		For Each oGallery In collGallery.Galleries.Items		
			lngGalleryId = oGallery.ID
			strGalleryLastName = oGallery.GalleryLastName
			strGalleryName = oGallery.GalleryName
			dtExpirationDate = oGallery.ExpirationDate
			intGalleryUser = oGallery.GalleryUserId
			intActiveFlag = oGallery.ActiveFlag
			Set oGallery = Nothing
		Next
		Set collGallery = Nothing
	End If
End If
%>
<HTML>
<HEAD>
<TITLE><%=PAGE_TITLE%></TITLE>
<script src="/cms/lib/__dom.js" type="text/javascript"></script>
<script src="/cms/lib/__calendar2.js" type="text/javascript"></script>
<script>
function checkForm() {
	var domUser = findDOM('txtUserId');
	if (domUser.options.value==0) {
		alert('You must associate a user to this gallery. Please select one from the drop down box, then try again.');
		return;
	}
	var frm = document.forms['form1'];

	if (!document.getElementById('txtGalleryLastName')) {
		var fn=frm.elements['txtGalleryLastName'].value;
		var gn=frm.elements['txtGalleryName'].value;
		var exp=frm.elements['txtExpirationDate'].value;	
	} else {
		var fn=document.getElementById('txtGalleryLastName').value;
		var gn=document.getElementById('txtGalleryName').value;
		var exp=document.getElementById('txtExpirationDate').value;	
	}

	if (fn=='') {
		alert('You must specify a unique proofing gallery name, which will be the folder name on the server.');
		return false;
	}
	if (gn=='') {
		alert('You must specify a proofing gallery description so you can easily identify the photos which reside in this gallery.');
		return false;
	}
	if (exp=='') {
		alert('You must specify an expiration date.');		
		return false;
	}

	var mode = frm.hidMode;
	mode.value = 'save';
	frm.submit();
}
</script>
<!--#include virtual="/cms/styles.asp"-->
</HEAD>
<BODY>
<form name="form1" id="form1" method="post" action="<%=EDIT_PAGE%>">
<input type="hidden" id="hidGalleryId" name="hidGalleryId" value="<%=lngGalleryId%>">
<input type="hidden" name="hidMode" id="hidMode" value="<%=strMode%>">
<input type="hidden" name="hidOldGalleryLastName" id="hidOldGalleryLastName" value="<%=strGalleryLastName%>">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" style="border-bottom-style:solid;border-width:1px;border-color:#666666;">
	<!--#include virtual="/cms/common/__header.asp"-->
	<!--#include virtual="/cms/common/__titlebar.asp" -->
	<!--#include virtual="/cms/common/__begin_bodywrap.asp"-->
	<table class="table-container" width="<%=CMS_PAGE_WIDTH%>">
		<tr>
			<th align="left" width="200">Client Full Name<em>*</em></th>
			<td align="left"><input type="text" size="50" id="txtGalleryLastName" name="txtGalleryLastName" value="<%=strGalleryLastName%>">&nbsp;<i>(make unique)</i></td>
		</tr>
		<tr>
			<th align="left" width="200">Gallery Identification<em>*</em></th>
			<td align="left"><input type="text" size="50" id="txtGalleryName" name="txtGalleryName" value="<%=strGalleryName%>">&nbsp;<i>(describe the type of photos)</i></td>
		</tr>
		<tr>
			<th align="left" width="200">Expiration Date<em>*</em></th>
			<td align="left"><input type="text" size="25" id="txtExpirationDate" name="txtExpirationDate" value="<%=dtExpirationDate%>">&nbsp;<i>(mm/dd/yyyy)</i>
<!--			<a href="javascript:;" onClick="calExpire.popup();"><img src="<%=CMS_IMAGE_PATH%>/calendar.gif" border="0" alt="Click to Pick a Expiration Date"></a>-->
			</td>			
		</tr>
		<%
		Response.write "<SCR" & "IPT>" & vbcrlf
		Response.write "var domExpire = findDOM('txtExpirationDate');" & vbcrlf
		Response.write "domExpire.value='" & dtExpirationDate & "';" & vbcrlf
		Response.write "domExpire.disabled=false;" & vbcrlf							
		Response.write "</SCR" & "IPT>" & vbcrlf
		%>				
		<tr>
			<th align="left">Gallery User<em>*</em></th>
			<td align="left"><%=GetUsersCombo("txtUserId",intGalleryUser)%></td>
		</tr>	
		<tr>
			<th align="left">Active?</th>
			<td align="left"><input type="checkbox" id="ckActiveFlag" name="ckActiveFlag" <%=IsChecked(intActiveFlag)%>></td>
		</tr>			
		<%=RenderRequiredFieldsMessageRow%>
	</table>
	<!--#include virtual="/cms/common/__end_bodywrap.asp"-->
</table>
<!--#include virtual="/cms/common/__footer.asp"-->
<script language="JavaScript">
<!-- // create calendar object(s) just after form tag closed
	var calExpire = new calendar2(domExpire);
	calExpire.year_scroll = true;
	calExpire.time_comp = false;
//-->
</script>
</form>

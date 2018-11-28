<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/cms/common/__fsoConfig.asp" -->
<!-- #include virtual="/cms/common/__fsoCommon.asp" -->
<!-- #include virtual="/objects/cCategory.asp" -->
<%
Dim PAGE_IMAGE, EDIT_PAGE, REPORT_PAGE, PAGE_TITLE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_content.gif"
PAGE_TITLE = "Portfolio Category"
EDIT_PAGE = "category.asp"
REPORT_PAGE = "category-report.asp"
IS_REPORT_PAGE = False
'ROOT_PATH = "/"
'PHOTO_IMAGE_PATH = ROOT_PATH & "secure/portfolio/" & PHOTOGRAPHER_FNAME & "/" 
Dim PHOTO_IMAGE_PATH
PHOTO_IMAGE_PATH = ROOT_PATH & GALLERY_PATH & "/"

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function AddCategoryText(CategoryText,CategoryCaption,pageOrder,intActiveFlag,lngCategoryId)

	Set fso = CreateObject("Scripting.FileSystemObject")
	
	'uncomment for localhost development
	'if Not fso.FolderExists(Server.MapPath(Replace(PHOTO_IMAGE_PATH & CategoryText & "/","/","\"))) Then
	'    fso.CreateFolder(GetFilePath(PHOTO_IMAGE_PATH & CategoryText))
	'    fso.CreateFolder(GetFilePath(PHOTO_IMAGE_PATH & CategoryText & "/thumbs"))
	
    'uncomment for production server
    Dim path,tpath
    path = Server.MapPath(Replace("/" & GALLERY_PATH & "/" & CategoryText & "/","/","\"))
    tpath = Server.MapPath(Replace("/" & GALLERY_PATH & "/" & CategoryText & "/thumbs","/","\"))
	
	if Not fso.FolderExists(path) Then
	    fso.CreateFolder(path)
	    fso.CreateFolder(tpath)	
	    AddCategoryText = True
	else
	    AddCategoryText = False
    end If
    Set fso = Nothing

	If AddCategoryText Then
	    Dim oCategory
	    Set oCategory = New cCategory
	    With oCategory
		    .CategoryText = CategoryText
		    .PageOrder = pageOrder
		    .CategoryCaption = CategoryCaption
		    .ActiveFlag = intActiveFlag
		    .AddCategoryText()
		    lngCategoryId = .ID
		    AddCategoryText = Not .IsError
	    End With
	    Set oCategory = Nothing    		
	End If
	
End Function

Function UpdateCategoryText(id,OldCategoryText,CategoryText,CategoryCaption,pageOrder,intActiveFlag)

	'If StrComp(OldCategoryText,CategoryText,0) <> 0 Then	

		'Set fso = CreateObject("Scripting.FileSystemObject")
		'if not fso.FolderExists(Server.MapPath(Replace("/" & GALLERY_PATH & "/" & CategoryText & "/","/","\"))) then

			Dim oCategory
			Set oCategory = New cCategory
			With oCategory
				.ID = id
				.CategoryText = CategoryText
				.CategoryCaption = CategoryCaption
				.PageOrder = pageOrder
				.ActiveFlag = intActiveFlag
				.UpdateCategoryText()
				UpdateCategoryText = Not .IsError
			End With
			Set oCategory = Nothing
	
		Set fso = CreateObject("Scripting.FileSystemObject")
		if not fso.FolderExists(Server.MapPath(Replace("/" & GALLERY_PATH & "/" & CategoryText & "/","/","\"))) then

			If UpdateCategoryText Then
				If StrComp(OldCategoryText,CategoryText,0) <> 0 Then	
					Set fso2 = CreateObject("Scripting.FileSystemObject")
					Set fldr2 = fso.GetFolder(Server.MapPath(Replace("/" & GALLERY_PATH & "/" & OldCategoryText & "/","/","\")))
					fldr2.Name = CategoryText
					Set fso2 = Nothing
					Set fldr2 = Nothing	
				End If
			End If
		'else
		'	UpdateCategoryText = False

		End If

		Set fso = Nothing	

	'Else
	'	UpdateCategoryText = True
	'End If

	
End Function


'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim lngCategoryId
Dim strOldCategoryText
Dim strCategoryText
Dim strCategoryCaption
Dim intPageOrder
Dim intActiveFlag

'----------------------------------------------------------------------------------------
' PAGE RENDER LOGIC
'----------------------------------------------------------------------------------------
If GetFormPost("hidMode") = "save" Then
	strMode = "edit"	
	lngCategoryId = GetFormPost("hidCategoryId")
	strOldCategoryText = GetFormPost("hidOldCategoryName")

	if StringEmptyOrNull(Session("OldCategoryText")) then
		Session("OldCategoryText") = strOldCategoryText & ""
	end if
	
	strCategoryText = GetFormPost("txtCategoryText")
	strCategoryCaption = QuoteCleanup(GetFormPost("taCategoryCaption"))
	intPageOrder = GetFormPost("txtPageOrder")

	'if StringEmptyOrNull(intPageOrder) Then
	'	intPageOrder = 1
	'End If

	intActiveFlag = GetSqlCheckboxValue(GetFormPost("ckActiveFlag"))
		
	If StringNotEmptyOrNull(lngCategoryId) Then
		blnSuccess = UpdateCategoryText(lngCategoryId,Session("OldCategoryText"),strCategoryText,strCategoryCaption,intPageOrder,intActiveFlag)		
	Else
		blnSuccess = AddCategoryText(strCategoryText,CategoryCaption,intPageOrder,intActiveFlag,lngCategoryId)
	End If
		
	If blnSuccess Then
		Session("OldCategoryText") = ""
		PageRedirect(REPORT_PAGE)
	Else
		displayMessage = "An error occurred while trying to save information to the database or file system. It could be that you were trying to create a category which already exists."
	End If
Else
    strMode = "save"
	lngCategoryId = GetQryString("id")
	If StringNotEmptyOrNull(lngCategoryId) Then
		Set oCategory = New cCategory
		Set collCategory = New cCategory
		collCategory.ID = lngCategoryId
		Call collCategory.GetCategoryTextById()
		For Each oCategory In collCategory.Categories.Items		
			lngCategoryId = oCategory.ID
			strCategoryText = oCategory.CategoryText
			strCategoryCaption = oCategory.CategoryCaption
			intPageOrder = oCategory.PageOrder
			intActiveFlag = oCategory.ActiveFlag
			Set oCategory = Nothing
		Next
		Set collCategory = Nothing
	End If
End If
%>
<HTML>
<HEAD>
<TITLE><%=PAGE_TITLE%></TITLE>
<script src="/cms/lib/__dom.js" type="text/javascript"></script>
<script>
function checkForm() {
	var frm = document.forms['form1'];

	if (!document.getElementById('txtPageOrder')) {
		var cn=frm.elements['txtCategoryText'].value;
		var po=frm.elements['txtPageOrder'].value;	
	} else {
		var cn=document.getElementById('txtCategoryText').value;
		var po=document.getElementById('txtPageOrder').value;	
	}
	if (cn=='') {
		alert('You must specify a category name');
		return false;
	}
	if (po=='') {
		alert('You must specify a page order.');		
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
<input type="hidden" id="hidCategoryId" name="hidCategoryId" value="<%=lngCategoryId%>">
<input type="hidden" name="hidOldCategoryName" id="hidOldCategoryName" value="<%=strCategoryText%>">
<input type="hidden" name="hidMode" id="hidMode" value="<%=strMode%>">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" style="border-bottom-style:solid;border-width:1;border-color:#666;">
	<!--#include virtual="/cms/common/__header.asp"-->
	<!--#include virtual="/cms/common/__titlebar.asp" -->
	<!--#include virtual="/cms/common/__begin_bodywrap.asp"-->
	<table class="table-container" width="<%=CMS_PAGE_WIDTH%>">
		<tr>
			<th align="left" width="200">Category Text<em>*</em></th>
			<td align="left"><input type="text" size="50" id="txtCategoryText" name="txtCategoryText" value="<%=strCategoryText%>"></td>
		</tr>
		<tr>
			<th align="left" width="200">Category Caption</th>
			<td align="left"><textarea cols="60" rows="2" class="admin-section" id="taCategoryCaption" name="taCategoryCaption"><%=strCategoryCaption%></textarea></td>
		</tr>
		<tr>
			<th align="left">Page Order<em>*</em></th>
			<td align="left"><input type="text" size="50" id="txtPageOrder" name="txtPageOrder" value="<%=intPageOrder%>"></td>
			</td>
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
</form>
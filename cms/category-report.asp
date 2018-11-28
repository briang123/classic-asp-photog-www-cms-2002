<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/cms/common/__fsoConfig.asp" -->
<!-- #include virtual="/cms/common/__fsoCommon.asp" -->
<!-- #include virtual="/objects/cCategory.asp" -->
<!-- #include virtual="/objects/cPhotos.asp" -->
<%
Dim PAGE_IMAGE, EDIT_PAGE, REPORT_PAGE, PAGE_TITLE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_content.gif"
PAGE_TITLE = "Portfolio Category"
EDIT_PAGE = "category.asp"
REPORT_PAGE = "category-report.asp"
IS_REPORT_PAGE = True
Dim PHOTO_IMAGE_PATH
PHOTO_IMAGE_PATH = ROOT_PATH & GALLERY_PATH & "/"

Dim path
path = replace(GetAppVariable("PHYSICAL_ROOT_PATH") & "/" & GetAppVariable("GALLERY_PATH") & "\","/","\")

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function DeleteCategory(id,categoryText)

	Dim success
	Dim oPhoto
	Set oPhoto = New cPhotos
	With oPhoto
		.CategoryId = id
		.DeletePhotosByCategory()
		success = Not .IsError
	End With
	
	If success Then
		Dim oCategory
		Set oCategory = New cCategory
		With oCategory
			.ID = id
			.DeleteCategoryText()
			DeleteCategory = Not .IsError
		End With
	
		If DeleteCategory Then
		    'uncomment for localhost development
			'objFSO.DeleteFolder(GetFilePath(PHOTO_IMAGE_PATH & categoryText))
			
			'uncomment for production server
			objFSO.DeleteFolder(path & categoryText)
			Set objFSO = Nothing
		Else
			DeleteCategory = False
			Exit Function
		End If
	Else
		DeleteCategory = False	
	End If
		
End Function


'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim intCategoryId

'----------------------------------------------------------------------------------------
' PAGE RENDNER LOGIC
'----------------------------------------------------------------------------------------
intCategoryId = GetQryString("id")
If intCategoryId > 0 Then
	blnSuccess = DeleteCategory(intCategoryId,GetQryString("cn"))
	
	If blnSuccess Then		
		displayMessage = "The information was successfully deleted from the system."
	Else
		displayMessage = "An error occurred while attempting to delete the information from the database."
	End If
End If
%>
<HTML>
<HEAD>
<TITLE><%=PAGE_TITLE%></TITLE>
<!--#include virtual="/cms/styles.asp"-->
</HEAD>
<BODY>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" style="border-bottom-style:solid;border-width:1;border-color:#666;">
	<!--#include virtual="/cms/common/__header.asp"-->
	<!--#include virtual="/cms/common/__titlebar.asp" -->
	<!--#include virtual="/cms/common/__begin_bodywrap.asp"-->
	<table class="report" width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td class="report-header">&nbsp;</td>
			<td class="report-header">Category Text</td>										
			<td class="report-header">Category Caption</td>
			<td class="report-header">Active Status</td>
			<td class="report-header">Order</td>
			<td class="report-header">&nbsp;</td>
		</tr>
		<%					
		Set oCategory = New cCategory
		Set collCategory = New cCategory
		Call collCategory.GetCategoryText()
		For Each oCategory In collCategory.Categories.Items		
		%>
			<tr>
				<td>
					<a href="<%=REPORT_PAGE%>?id=<%=oCategory.ID%>&cn=<%=oCategory.CategoryText%>"><img src="<%=CMS_IMAGE_PATH%>/delitem.gif" vspace="10" alt="Delete"></a>&nbsp;
					<a href="<%=EDIT_PAGE%>?id=<%=oCategory.ID%>"><img src="<%=CMS_IMAGE_PATH%>/edit.gif" vspace="10" alt="Edit"></a>
				</td>
				<td><%If Len(oCategory.CategoryText) > 0 Then 
						echo(oCategory.CategoryText)
					Else 
						echo("(empty)")
					End If%></td>	
				<td><%If Len(oCategory.CategoryCaption) > 0 Then 
						echo(oCategory.CategoryCaption)
					Else 
						echo("(empty)")
					End If%></td>
				<td><%=GetCheckboxValue(oCategory.ActiveFlag,True)%></td>									
				<td><%=oCategory.PageOrder%></td>
				<td>
					<a href="photo-report.asp?cid=<%=oCategory.ID%>&cn=<%=oCategory.CategoryText%>"><img src="<%=CMS_IMAGE_PATH%>/photolist.gif" alt="View all connected photos for this category"></a>&nbsp;
					<a href="photo-workspace.asp?cid=<%=oCategory.ID%>&cn=<%=oCategory.CategoryText%>"><img src="<%=CMS_IMAGE_PATH%>/addimage.gif" alt="Upload/Link Photos for this category"></a>
				</td>				
			</tr>									
		<%
			Set oCategory = Nothing
		Next
		Set collCategory = Nothing
		%>
	</table>																	
	<!--#include virtual="/cms/common/__end_bodywrap.asp"-->
</table>
<!--#include virtual="/cms/common/__footer.asp"-->
</BODY>
</HTML>

<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/cms/common/__fsoConfig.asp" -->
<!-- #include virtual="/cms/common/__fsoCommon.asp" -->
<!-- #include virtual="/objects/cGallery.asp" -->
<!-- #include virtual="/objects/cPhotos.asp" -->
<%
Dim PAGE_IMAGE, EDIT_PAGE, REPORT_PAGE, PAGE_TITLE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_desfield.gif"
PAGE_TITLE = "Proofing Gallery"
EDIT_PAGE = "gallery.asp"
REPORT_PAGE = "gallery-report.asp"
IS_REPORT_PAGE = True

'Dim PHOTO_IMAGE_PATH
'ROOT_PATH = "/"
'PHOTO_IMAGE_PATH = ROOT_PATH & "secure/proofs/" & PHOTOGRAPHER_FNAME & "/" 
Dim PHOTO_IMAGE_PATH
'PHOTO_IMAGE_PATH = CMS_ROOT_PATH & PROOF_PATH & "/"
PHOTO_IMAGE_PATH = ROOT_PATH & PROOF_PATH & "/"

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function DeleteGallery(id,galleryLastName)

	Dim success
	Dim oPhoto
	Set oPhoto = New cPhotos
	With oPhoto
		.GalleryId = id
		.DeletePhotosByGallery()
		success = Not .IsError
	End With
	
	If success Then
		Dim oGallery
		Set oGallery = New cGallery
		With oGallery
			.ID = id
			.DeleteGallery()
			DeleteGallery = Not .IsError
		End With
	
		If DeleteGallery Then
		    'uncomment localhost development
			'objFSO.DeleteFolder(GetFilePath(PHOTO_IMAGE_PATH & LCase(galleryLastName)))
			
			'uncomment production server
			objFSO.DeleteFolder(GetFilePath("/" & PROOF_PATH & "/" & LCase(galleryLastName)))
			Set objFSO = Nothing
		Else
			DeleteGallery = False
			Exit Function
		End If
	Else
		DeleteGallery = False	
	End If

End Function

'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim intGalleryId

'----------------------------------------------------------------------------------------
' PAGE RENDNER LOGIC
'----------------------------------------------------------------------------------------
intGalleryId = GetQryString("id")
If intGalleryId > 0 Then
	blnSuccess = DeleteGallery(intGalleryId,GetQryString("gln"))
	
	If blnSuccess Then		
		displayMessage = "The gallery and all associated photographs were deleted from the server and database."
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
			<td class="report-header">Client Last Name</td>
			<td class="report-header">Gallery Name</td>
			<td class="report-header">Expiration Date</td>
			<td class="report-header">Gallery User</td>										
			<td class="report-header">Active Status</td>
			<td class="report-header">&nbsp;</td>
		</tr>
		<%					
		Set oGallery = New cGallery
		Set collGallery = New cGallery
		Call collGallery.GetGallery()
		For Each oGallery In collGallery.Galleries.Items		
		%>
			<tr>
				<td>
					<a href="<%=REPORT_PAGE%>?id=<%=oGallery.ID%>&gln=<%=LCase(oGallery.GalleryLastName)%>"><img src="<%=CMS_IMAGE_PATH%>/delitem.gif" vspace="10" alt="Delete"></a>&nbsp;
					<a href="<%=EDIT_PAGE%>?id=<%=oGallery.ID%>"><img src="<%=CMS_IMAGE_PATH%>/edit.gif" vspace="10" alt="Edit"></a>
				</td>
				<td><%=oGallery.GalleryLastName%></td>
				<td><%=oGallery.GalleryName%></td>
				<td><%=oGallery.ExpirationDate%></td>
				<td><%=oGallery.GalleryUser%></td>				
				<td><%=GetCheckboxValue(oGallery.ActiveFlag,True)%></td>
				<td>
				<a href="photo-report.asp?gid=<%=oGallery.ID%>&gln=<%=oGallery.GalleryLastName%>&gn=<%=oGallery.GalleryName%>"><img src="<%=CMS_IMAGE_PATH%>/photolist.gif" alt="View all connected photos for this gallery"></a>&nbsp;
				<a href="photo-workspace.asp?gid=<%=oGallery.ID%>&gln=<%=oGallery.GalleryLastName%>&gn=<%=oGallery.GalleryName%>"><img src="<%=CMS_IMAGE_PATH%>/addimage.gif" alt="Upload/Link Photos for this gallery"></a></td>
			</tr>									
		<%
			Set oGallery = Nothing
		Next
		Set collGallery = Nothing
		%>
	</table>																	
	<!--#include virtual="/cms/common/__end_bodywrap.asp"-->
</table>
<!--#include virtual="/cms/common/__footer.asp"-->
</BODY>
</HTML>

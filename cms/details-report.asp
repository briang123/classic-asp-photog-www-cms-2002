<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/objects/cDetails.asp" -->
<%
Dim PAGE_IMAGE, EDIT_PAGE, REPORT_PAGE, PAGE_TITLE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_content.gif"
PAGE_TITLE = "Session Details"
EDIT_PAGE = "details.asp"
REPORT_PAGE = "details-report.asp"
IS_REPORT_PAGE = True

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function DeleteDetails(id)
	Dim oDetails
	Set oDetails = New cDetails
	With oDetails
		.ID = id
		.DeleteDetails()
		DeleteDetails = Not .IsError
	End With
End Function

'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim intDetailsId

'----------------------------------------------------------------------------------------
' PAGE RENDNER LOGIC
'----------------------------------------------------------------------------------------
intDetailsId = GetQryString("id")
If intDetailsId > 0 Then
	blnSuccess = DeleteDetails(intDetailsId)
	
	If blnSuccess Then		
		displayMessage = "The deletion was success."
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
			<td class="report-header">Description</td>										
		</tr>
		<%					
		Set oDetails = New cDetails
		Set collDetails = New cDetails
		Call collDetails.GetDetails()
		For Each oDetails In collDetails.Details.Items		
		%>
			<tr>
				<td>
					<a href="<%=REPORT_PAGE%>&id=<%=oDetails.ID%>"><img src="<%=CMS_IMAGE_PATH%>/delitem.gif" vspace="10" alt="Delete"></a>&nbsp;
					<a href="<%=EDIT_PAGE%>&id=<%=oDetails.ID%>"><img src="<%=CMS_IMAGE_PATH%>/edit.gif" vspace="10" alt="Edit"></a>
				</td>
				<td><%echo(oDetails.DetailsText)%></td>											
			</tr>									
		<%
			Set oDetails = Nothing
		Next
		Set collDetails = Nothing
		%>
	</table>																	
	<!--#include virtual="/cms/common/__end_bodywrap.asp"-->
</table>
<!--#include virtual="/cms/common/__footer.asp"-->
</BODY>
</HTML>

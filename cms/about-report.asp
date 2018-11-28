<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/objects/cAbout.asp" -->
<%
Dim PAGE_IMAGE, EDIT_PAGE, REPORT_PAGE, PAGE_TITLE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_content.gif"
PAGE_TITLE = "About Me"
EDIT_PAGE = "about.asp"
REPORT_PAGE = "about-report.asp"
IS_REPORT_PAGE = True

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function DeleteAboutText(id)
	Dim oAbout
	Set oAbout = New cAbout
	With oAbout
		.ID = id
		.DeleteAboutText()
		DeleteAboutText = Not .IsError
	End With
End Function

'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim intAboutId

'----------------------------------------------------------------------------------------
' PAGE RENDNER LOGIC
'----------------------------------------------------------------------------------------
intAboutId = GetQryString("id")
If intAboutId > 0 Then
	blnSuccess = DeleteAboutText(intAboutId)
	
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
		Set oAbout = New cAbout
		Set collAbout = New cAbout
		Call collAbout.GetAboutText()
		For Each oAbout In collAbout.Abouts.Items		
		%>
		<tr>
			<td>
				<a href="<%=REPORT_PAGE%>&id=<%=oAbout.ID%>"><img src="<%=CMS_IMAGE_PATH%>/delitem.gif" vspace="10" alt="Delete"></a>&nbsp;
				<a href="<%=EDIT_PAGE%>&id=<%=oAbout.ID%>"><img src="<%=CMS_IMAGE_PATH%>/edit.gif" vspace="10" alt="Edit"></a>				
			</td>
			<td><%echo(oAbout.AboutText)%></td>
		</tr>									
		<%
			Set oAbout = Nothing
		Next
		Set collAbout = Nothing
		%>
	</table>																	
	<!--#include virtual="/cms/common/__end_bodywrap.asp"-->
</table>
<!--#include virtual="/cms/common/__footer.asp"-->
</BODY>
</HTML>

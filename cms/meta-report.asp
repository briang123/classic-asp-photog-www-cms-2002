<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/objects/cMetaData.asp" -->
<%
Dim PAGE_IMAGE, PAGE_TITLE, EDIT_PAGE, REPORT_PAGE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_listset.gif"
PAGE_TITLE = "Meta Tag Report"
EDIT_PAGE = "meta.asp"
REPORT_PAGE = "meta-report.asp"
IS_REPORT_PAGE = True

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function DeleteMetaData(id)
	Dim oMeta
	Set oMeta = New cMetaData
	With oMeta
		.ID = id
		.DeleteMetaData()
		DeleteMetaData = Not .IsError
	End With
End Function

'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim intMetaId

'----------------------------------------------------------------------------------------
' PAGE RENDNER LOGIC
'----------------------------------------------------------------------------------------
intMetaId = GetQryString("id")
If intMetaId > 0 Then
	blnSuccess = DeleteMetaData(intMetaId)
	
	If blnSuccess Then		
		displayMessage = "The deletion of the meta tag information was successful."
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
	<% CMS_PAGE_WIDTH = "100%" %>
	<!--#include virtual="/cms/common/__begin_bodywrap.asp"-->
	<table class="report" width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td class="report-header">&nbsp;</td>
			<td class="report-header">Web Page</td>
			<td class="report-header">MetaTag Keywords</td>
			<td class="report-header">MetaTag Description</td>
		</tr>
		<%			
		Set oMeta = New cMetaData
		Set collMeta = New cMetaData
		Call collMeta.GetMetaData()
		For Each oMeta In collMeta.MetaData.Items
		%>
			<tr>
				<td width="50">
					<a href="<%=REPORT_PAGE%>?id=<%=oMeta.ID%>"><img src="<%=CMS_IMAGE_PATH%>/delitem.gif" vspace="10" alt="Delete"></a>&nbsp;
					<a href="<%=EDIT_PAGE%>?id=<%=oMeta.ID%>"><img src="<%=CMS_IMAGE_PATH%>/edit.gif" vspace="10" alt="Edit"></a>
				</td>
				<td width="100"><%=oMeta.WebPage%></td>
				<td width="250"><%=oMeta.MetaKeywords%>&nbsp;</td>	
				<td><%=oMeta.MetaDescription%>&nbsp;</td>										
			</tr>									
		<%
				Set oMeta = Nothing
		Next
		Set collMeta = Nothing
		%>
	</table>																	
	<!--#include virtual="/cms/common/__end_bodywrap.asp"-->
</table>
<!--#include virtual="/cms/common/__footer.asp"-->
</BODY>
</HTML>

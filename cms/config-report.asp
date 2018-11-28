<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<%
Dim PAGE_IMAGE, PAGE_TITLE, EDIT_PAGE, REPORT_PAGE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_listset.gif"
PAGE_TITLE = "Configuration Management Report"
EDIT_PAGE = "config.asp"
REPORT_PAGE = "config-report.asp"
IS_REPORT_PAGE = True

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function DeleteConfig(id)
	Dim oConfig
	Set oConfig = New cConfig
	With oConfig
		.ID = id
		.DeleteConfig()
		DeleteConfig = Not .IsError
	End With
End Function

'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim intConfigId

'----------------------------------------------------------------------------------------
' PAGE RENDNER LOGIC
'----------------------------------------------------------------------------------------
intConfigId = GetQryString("id")
If intConfigId > 0 Then
	blnSuccess = DeleteConfig(intConfigId)
	
	If blnSuccess Then		
		On Error Resume Next
		Dim strKey
		strKey = GetQryString("key")
		Application.Contents.Remove(strKey)
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
<% CMS_PAGE_WIDTH = "100%" %>
	<!--#include virtual="/cms/common/__begin_bodywrap.asp"-->
	<table class="report" width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td class="report-header" width="50px">&nbsp;</td>
			<td class="report-header">Config Key</td>
			<td class="report-header">Config Value</td>
			<td class="report-header" width="200">Config Description</td>
		</tr>
		<%			
		Set oConfig = New cConfig
		Set collConfig = New cConfig
		Call collConfig.GetConfigs()
		For Each oConfig In collConfig.Configs.Items
		%>
			<tr>
				<td>
					<a href="<%=REPORT_PAGE%>?id=<%=oConfig.ID%>&key=<%=oConfig.ConfigKey%>"><img src="<%=CMS_IMAGE_PATH%>/delitem.gif" vspace="10" alt="Delete"></a>&nbsp;
					<a href="<%=EDIT_PAGE%>?id=<%=oConfig.ID%>"><img src="<%=CMS_IMAGE_PATH%>/edit.gif" vspace="10" alt="Edit"></a>
				</td>
				<td><%=oConfig.ConfigKey%></td>
				<td><%If Len(oConfig.ConfigValue) > 0 Then 
						echo(oConfig.ConfigValue)
					Else 
						echo("(empty)")
					End If%></td>	
				<td><%If Len(oConfig.ConfigDesc) > 0 Then 
						echo(oConfig.ConfigDesc)
					Else 
						echo("(empty)")
					End If%></td>										
			</tr>									
		<%
				Set oConfig = Nothing
		Next
		Set collConfig = Nothing
		%>
	</table>																	
	<!--#include virtual="/cms/common/__end_bodywrap.asp"-->
</table>
<!--#include virtual="/cms/common/__footer.asp"-->
</BODY>
</HTML>

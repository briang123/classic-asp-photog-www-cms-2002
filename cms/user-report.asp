<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/objects/cLogin.asp" -->
<%
Dim PAGE_IMAGE, EDIT_PAGE, REPORT_PAGE, PAGE_TITLE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_content.gif"
PAGE_TITLE = "User List"
EDIT_PAGE = "user.asp"
REPORT_PAGE = "user-report.asp"
IS_REPORT_PAGE = True

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function DeleteUser(id)
	Dim oLogin
	Set oLogin = New cLogin
	With oLogin
		.ID = id
		.DeleteUser()
		DeleteUser = Not .IsError
	End With
End Function


'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim intUserId

'----------------------------------------------------------------------------------------
' PAGE RENDNER LOGIC
'----------------------------------------------------------------------------------------
intUserId = GetQryString("id")
If intUserId > 0 Then
	blnSuccess = DeleteUser(intUserId)
	
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
	<% CMS_PAGE_WIDTH = "750" %>	
	<!--#include virtual="/cms/common/__begin_bodywrap.asp"-->
	<table class="report" width="750" cellpadding="0" cellspacing="0">
		<tr>
			<td class="report-header">&nbsp;</td>
			<td class="report-header">Full Name</td>
			<td class="report-header">Login Info</td>
			<!--<td class="report-header">Email Address</td>
			<td class="report-header">Address Info</td>
			<td class="report-header">Comments</td>-->
			<td class="report-header">Expire</td>
			<td class="report-header">CMS Admin</td>
		</tr>
		<%					
		Set oLogin = New cLogin
		Set collLogin = New cLogin
		Call collLogin.GetUsers()
		For Each oLogin In collLogin.Logins.Items		
		%>
			<tr>
				<td>
					<a href="<%=REPORT_PAGE%>?id=<%=oLogin.ID%>"><img src="<%=CMS_IMAGE_PATH%>/delitem.gif" vspace="10" alt="Delete"></a>&nbsp;
					<a href="<%=EDIT_PAGE%>?id=<%=oLogin.ID%>"><img src="<%=CMS_IMAGE_PATH%>/edit.gif" vspace="10" alt="Edit"></a>
				</td>
				<td><%=oLogin.FullName%></td>
				<td><% 
					echobr("login: " & oLogin.Login)
					echobr("password: " & oLogin.Password) 
				%></td>
<!--				
                <td><%=oLogin.Email%></td>
				<td><% 
					echobr(oLogin.Address1)
					if StringNotEmptyOrNull(oLogin.Address2) Then echobr(oLogin.Address2)
					echobr(oLogin.City & ", " & oLogin.StateCode & " " & oLogin.Zip)
					echo(oLogin.Phone)
					%></td>
				<td width="300"><% If StringNotEmptyOrNull(oLogin.Comments) Then echo(oLogin.Comments) Else echo("&nbsp;")%></td>
-->
				<td><%=oLogin.Expire%></td>
				<td><%=oLogin.IsAdmin %></td>
			</tr>									
		<%
			Set oLogin = Nothing
		Next
		Set collLogin = Nothing
		%>
	</table>																	
	<!--#include virtual="/cms/common/__end_bodywrap.asp"-->
</table>
<!--#include virtual="/cms/common/__footer.asp"-->
</BODY>
</HTML>

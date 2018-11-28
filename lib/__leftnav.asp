<%
REsponse.write "Side Navigation<br>"
if 1 = 0 then

Dim rsLeftNav
Dim intSubMenuId
Dim intSubSubMenuId
Dim strSqlLeftNav
Dim strLeftNavTableColor
Dim strUrl
Dim strLeftNavMenuName
Dim strMouseOverColor
%>
<table border="0" cellpadding="0" cellspacing="0">
<tr valign="top">
	<td colspan="2" align="center" valign="top">
<%
With Response
	.Write Session(APPVARNAME & "__UserFullName") & "<br/>"
	.Write Session(APPVARNAME & "__company") & "<br/>"
End With
%>
<b><a href="./login.asp?logout=1&user=<%=Server.URLEncode(session("__Email"))%>"><img src="<%=Application(APPVARNAME & "AppImagePath")%>/unlock.gif" border="0" alt="Logout">&nbsp;LOGOUT</a></b>
	</td>
</tr>
<% If Session(APPVARNAME & "Role") <> "Reader" Then %>
<tr><td colspan="2" valign="top">&nbsp;</td></tr>
<tr><td bgcolor="<%=strLeftNavTableColor%>" colspan="2" align="center" valign="top" class="<%=strLeftNavHeaderClassName%>">Request Management</td></tr>
<tr>
	<td onMouseover=this.bgColor="<%=strMouseOverColor%>" onMouseout=this.bgColor="<%=strLeftNavColor%>" style="cursor:hand" colspan="2">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td width="20px"><img src="<%=Application(APPVARNAME & "AppImagePath")%>/update.gif" border="0" alt="Add Request"></td>
				<td align="left"><a href="./reqadmin.asp?page=0">Add Request</a></td>
			</tr>
		</table>
	</td>
</tr>
<tr>
	<td onMouseover=this.bgColor="<%=strMouseOverColor%>" onMouseout=this.bgColor="<%=strLeftNavColor%>" style="cursor:hand" colspan="2">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td width="20px"><img src="<%=Application(APPVARNAME & "AppImagePath")%>/edititem.gif" border="0" alt="Modify Request"></td>
				<td align="left"><a href="./requests.asp?page=1">Modify Request</a></td>
			</tr>
		</table>
	</td>
</tr>
<% 	if Session(APPVARNAME & "Role") = "Master" Or Session(APPVARNAME & "Role") = "Administrator" then %>
<tr>
	<td onMouseover=this.bgColor="<%=strMouseOverColor%>" onMouseout=this.bgColor="<%=strLeftNavColor%>" style="cursor:hand" colspan="2">
		<table width="100%" border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td width="20px"><img src="<%=Application(APPVARNAME & "AppImagePath")%>/delete.gif" border="0" alt="Remove Request"></td>
				<td align="left"><a href="./requests.asp?page=2">Remove Request</a></td>
			</tr>
		</table>
	</td>
</tr>
<% 	end if 
End If
%>
<tr valign="top"><td valign="top">&nbsp;</td></tr>
<%
Call OpenRS(rsLeftNav)

'strSqlLeftNav = "SELECT DISTINCT(s.status), s.statusId, r.active " & _
'								"FROM tblStatus s, tblRequests r, tblProjects p " & _
'								"WHERE r.status = s.statusId " & _
'								"AND r.projId = " & session("ProjectId") & _
'								" AND r.active = True" & _
'								" ORDER BY s.statusid ASC"
								
'rsLeftNav.Open strSqlLeftNav, GetConnection, adOpenDynamic

With rsLeftNav
	If Not .BOF And Not .EOF Then
%>
	<tr valign="top"><td bgcolor="<%=strLeftNavTableColor%>" colspan="2" align="center" valign="top" class="<%=strLeftNavHeaderClassName%>">Status Reports</td></tr>
	<tr valign="top">
		<td onMouseover=this.bgColor="<%=strMouseOverColor%>" onMouseout=this.bgColor="<%=strLeftNavColor%>" style="cursor:hand" colspan="2"><a href="./requests.asp?search=No&status=0">All Requests</a></td>
	</tr>
<%
		.MoveFirst
		Do While Not .EOF
%>
	<tr>
		<td onMouseover=this.bgColor="<%=strMouseOverColor%>" onMouseout=this.bgColor="<%=strLeftNavColor%>" style="cursor:hand" colspan="2"><a href="./requests.asp?search=No&status=<%=rsLeftNav("statusId")%>"><%=rsLeftNav("status")%> Requests</a></td>
	</tr>
<%
			.MoveNext
		Loop		

		If Session(APPVARNAME & "Role") <> "Reader" Then
%>
	<tr><td colspan="2">&nbsp;</td></tr>
	<tr><td bgcolor="<%=strLeftNavColor%>" colspan="2"><img src="<%=Application(APPVARNAME & "AppImagePath")%>/members.gif" border="0">&nbsp;Personalized Reports</td></tr>	
	<tr><td bgcolor="<%=strLeftNavTableColor%>" colspan="2" align="center" class="<%=strLeftNavHeaderClassName%>"><b>Requests Opened</b></td></tr>
	<tr>
		<td onMouseover=this.bgColor="<%=strMouseOverColor%>" onMouseout=this.bgColor="<%=strLeftNavColor%>" style="cursor:hand" colspan="2"><a href="./requests.asp?search=No&status=0&user=1">All Requests</a></td>
	</tr>
<%
			.MoveFirst
			Do While Not .EOF
%>
	<tr>
		<td onMouseover=this.bgColor="<%=strMouseOverColor%>" onMouseout=this.bgColor="<%=strLeftNavColor%>" style="cursor:hand" colspan="2"><a href="./requests.asp?search=No&status=<%=rsLeftNav("statusId")%>&user=1"><%=rsLeftNav("status")%> Requests</a></td>
	</tr>
<%
				.MoveNext
			Loop
		End If
	End If
End With

Call CloseRS(rsLeftNav)

if Session(APPVARNAME & "Role") = "Master" then 
%>
<tr><td colspan="2">&nbsp;</td></tr>
<tr><td bgcolor="<%=strAdminHeaderColor%>" colspan="2" align="center" class="<%=strLeftNavHeaderClassName%>">Administration</td></tr>
<tr><td bgcolor="<%=strLeftNavColor%>" colspan="2"><a href="./config.asp">Configuration</a></td></tr>
<tr>
	<td bgcolor="<%=strLeftNavColor%>" colspan="2">
		<a class="menuHead" href="javascript:toggleAdminMenu('menu1')"><img src="<%=Application(APPVARNAME & "AppImagePath")%>/plus.gif" border="0">&nbsp;Projects</a>
	</td>
</tr>
<tr><td colspan="2" class="SubMenuIndent">
	<span id="menu1">
	<table width="100%" cellpadding="0" cellspacing="0" align="left">
		<tr><td><a href="./project.asp?page=2">Add New</a></td></tr>
		<tr><td><a href="./projectlist.asp">Modify</a></td></tr>
	</table>
	</span>		
</td>
</tr>
<tr>
	<td bgcolor="<%=strLeftNavColor%>" colspan="2">
		<a class="menuHead" href="javascript:toggleAdminMenu('menu3')"><img src="<%=Application(APPVARNAME & "AppImagePath")%>/plus.gif" border="0">&nbsp;Status</a>
	</td>
</tr>
<tr>
	<td bgcolor="<%=strLeftNavColor%>" colspan="2" class="SubMenuIndent">
		<span id="menu3">	
		<table width="100%" cellspacing="1" cellpadding="0" align="left">
			<tr><td><a href="./status.asp?page=2">Add New</a></td></tr>
			<tr><td><a href="./statuslist.asp">Modify</a></td></tr>
		</table>
		</span>	
	</td>
</tr>
<%	end if %>
<tr>
	<td bgcolor="<%=strLeftNavColor%>" colspan="2">
		<a class="menuHead" href="javascript:toggleAdminMenu('menu2')"><img src="<%=Application(APPVARNAME & "AppImagePath")%>/plus.gif" border="0">&nbsp;Users</a>	
	</td>
</tr>
<tr><td bgcolor="<%=strLeftNavColor%>" colspan="2" class="SubMenuIndent">
	<span id="menu2">	
	<table width="100%" cellspacing="1" cellpadding="0" align="left">
		<tr><td><a href="./users.asp?page=2">Add New</a></td></tr>
		<tr><td><a href="./userlist.asp">Modify</a></td></tr>
	</table>
	</span>	
</td>
</tr>
<tr><td bgcolor="<%=strLeftNavColor%>" colspan="2"><a href="./details.asp">Request Details</a></td></tr>
<tr><td colspan="2">&nbsp;</td></tr>
<tr><td bgcolor="<%=strAdminHeaderColor%>" colspan="2" align="center" class="<%=strLeftNavHeaderClassName%>">Admin Reports</td></tr>
<tr><td bgcolor="<%=strLeftNavColor%>" colspan="2"><a href="./requests.asp?search=No&page=0&inactive=1">Inactive Requests</a></td></tr>
<tr><td bgcolor="<%=strLeftNavColor%>" colspan="2"><a href="./requests.asp?search=No&page=0&inactive=2">Active Requests</a></td></tr>
</table>
<style>
<!-- 
#menu1 {display: none;}
#menu2 {display: none;}
#menu3 {display: none;}
-->
</style>
<% end if %>
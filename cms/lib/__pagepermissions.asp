<%
If CBool(Instr(url,"cms/")) Then

	Dim IsAdminUser, UserFullName
	
	IsAdminUser = CBool(GetSessionVariable("IS_ADMIN_USER"))
	UserFullName = LCase(GetSessionVariable("FULLNAME"))
	
	If 	(IsAdminUser = False) Then

		PageRedirect("/")

	End If
End If	
%>

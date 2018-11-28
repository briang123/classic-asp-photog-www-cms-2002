<%
If Not Session("COMPANY_SESS_IS_AUTHENTICATED") = True Then
	Response.Redirect("login.asp")
End If

Sub CheckUserAuthentication(strRedirect)
	
	' Redirect the user to the login page if not authenticated
	if Not GetSessionVariable(APPVARNAME & "UserAuthenticated") = True Then 
		Response.Redirect strRedirect
	End If
	
	' Redirect if the application is locked out except if Master user
	If GetAppVariable(APPVARNAME & "AppLockApp") = True Then
		If GetSessionVariable(APPVARNAME & "Role") <> "Master" Then
			Response.Redirect strRedirect
		End If
	End If

End Sub

'this will change
Function HasPermissions(strPage)
'Access Level Heirarchy
'1. Master - Add/Edit/Delete/Read all pages
'2. Administrator - Add/Edit/Read website pages, some administration pages
'3. Editor - Add/Edit/Read website pages
'4. Reader - Read only website pages

	Dim strUserRoleName
	If GetSessionVariable("__UserId") = 1 Then
		strUserRoleName = "DEVELOPER"
	End If
	
	Dim strPermissionLevel
	Select Case UCase(strPage)
		Case "GLOBALS"
			strPermissionLevel = "DEVELOPER"
		Case "USERS"
			strPermissionLevel = "DEVELOPER,MASTER,ADMINISTRATOR"
	End Select
	
	If strPermissionLevel = strUserRoleName Then
		HasPermissions = True
	Else
		HasPermissions = False
	End If
		
End Function
%>
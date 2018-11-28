<%
' =========================================================================
' CREATED BY: BRIAN GAINES
' FILENAME:		__appvarupdate.asp					
'	PURPOSE:		Load cache with application related information
' USAGE: 			Included when user authenticated
' =========================================================================

If Not AppVarsLoaded("CONFIGINFO") Then
	Call LoadConfigInfo(APPVARNAME, 1)
End If

If Not AppVarsLoaded("COMPANYINFO") Then
	LoadCompanyInfo
End If
%>

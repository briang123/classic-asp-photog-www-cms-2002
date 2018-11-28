<!-- #include virtual="/company/cms/lib/cCMSPagePermission.asp" -->
<%
'----------------------------------------------------------------------------------------
' NOTE:
'	For CMS pages which entail developer involvement, a filter has been set on the 
'	report pages so that no one, but someone within the developer's group can delete 
'	the entry. In some regards, the entire entry on the report page has been removed 
'	for viewing by anyone other than a developer. This holds true for configuration 
'	settings.
'
'	The following .asp files may need to be modified if working with permissions:
'		cCMSPagePermission.asp -- class file that talks to database
'		__pagepermissions.asp -- file which handles accessibility of pages
'		__global.asp -- has function which handles reports edit/delete controls
'		__toolbar.asp -- has functionality which handles toolbar controls
'		sidenav.asp -- file which handles rendering of page links
'		*.asp - other files which may have logic directly in them specific to page				
'----------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------
' ERROR PAGES
'----------------------------------------------------------------------------------------
Dim CMS_ACCESS_DENIED_ERROR_PAGE
CMS_ACCESS_DENIED_ERROR_PAGE = "accessdenied.asp"

'----------------------------------------------------------------------------------------
' COMMON FUNCTIONS << USED ON THIS PAGE ... AND ELSEWHERE
'----------------------------------------------------------------------------------------
Function HasGroupPermission(groupToCompare)
	If GetSessionVariable("USER_GROUP_LIST") = "" Then
		HasGroupPermission = False
	Else
		Dim groups, i
		HasGroupPermission = False
		groups = Split(GetSessionVariable("USER_GROUP_LIST"),",")
		For i = LBound(groups) To UBound(groups)
			HasGroupPermission = (CInt(groupToCompare) = CInt(groups(i)))
			If HasGroupPermission Then Exit Function
		Next 
	End If
End Function

Function HasCurrentPermission(permId)
	If StringNotEmptyOrNull(permId) Then
		HasCurrentPermission = CBool(Instr(Replace(permId,",",""),"1"))
	Else
		HasCurrentPermission = False
	End If
End Function

Function HasGroupPermission(groupToCompare)
	If GetSessionVariable("USER_GROUP_LIST") = "" Then
		HasGroupPermission = False
	Else
		Dim groups, i
		HasGroupPermission = False
		groups = Split(GetSessionVariable("USER_GROUP_LIST"),",")
		For i = LBound(groups) To UBound(groups)
			HasGroupPermission = (CInt(groupToCompare) = CInt(groups(i)))
			If HasGroupPermission Then Exit Function
		Next 
	End If
End Function

Function IsPage(page)
	IsPage = CBool(Instr(LCase(Request.ServerVariables("PATH_TRANSLATED")),LCase(page)))
End Function

Function HasPermissionWhileCMSIsLocked()
	HasPermissionWhileCMSIsLocked = Eval(HasGroupPermission(DEVELOPER_GRP_ID) Or HasGroupPermission(MASTER_GRP_ID))
End Function

Function GetGroupPermissionId()
	'Check if developer/master
	If HasGroupPermission(DEVELOPER_GRP_ID) Then
		GetGroupPermissionId = DEVELOPER_GRP_ID
	ElseIf HasGroupPermission(MASTER_GRP_ID) Then
		GetGroupPermissionId = MASTER_GRP_ID
	Else
		GetGroupPermissionId = 0
	End If
End Function

'----------------------------------------------------------------------------------------
' GROUP CONSTANTS - SHOULD NOT DELETE FROM DATABASE (IF DELETED AND RE-ADDED, MAKE SURE
'					THE IDEAS JIVE WITH DATABASE IDS; OTHERWISE IT WILL BREAK;
'					IF NEW GROUPS ARE ADDED, THEN ADD A CONSTANT HERE FOR IT AND APPLY
'					PERMISSIONS AS USUAL. KEEP IN MIND THAT ANY GROUP ADDED TO THE 
'					CMS WILL RESULT IN CODE CHANGES TO HANDLE WHAT TO DO WITH THAT GROUP.
'----------------------------------------------------------------------------------------
CONST DEVELOPER_GRP = "DEVELOPER"
CONST MASTER_GRP = "MASTER"
CONST MKT_OWNER_GRP = "MARKET OWNER"
CONST EDITOR_GRP = "CONTENT EDITOR"
CONST BUS_OWNER_GRP = "BUSINESS OWNER"
CONST DEVELOPER_GRP_ID = 1
CONST MASTER_GRP_ID = 2
CONST MKT_OWNER_ID = 3
CONST EDITOR_GRP_ID = 4
CONST BUS_OWNER_GRP_ID = 5

'----------------------------------------------------------------------------------------
' SPECIAL PAGES WHICH DO NOT HAVE A EDIT/REPORT PAGE SEPARATION
'----------------------------------------------------------------------------------------
Dim defaultPage
Dim loginPage
Dim denyPage
Dim profilePage

defaultPage = IsPage("default")
loginPage = IsPage("login")
denyPage = IsPage(CMS_ACCESS_DENIED_ERROR_PAGE)
profilePage = IsPage("profile")

'----------------------------------------------------------------------------------------
' CMS PAGES WHICH ONLY SOMEONE IN THE DEVELOPER OR MASTER GROUP SHOULD VIEW
'----------------------------------------------------------------------------------------
Dim configPage
Dim themePage
Dim userPage
Dim cmsPageTypePage
Dim groupPage
Dim cmsPagePage

configPage = IsPage("config")
themePage = IsPage("theme")
cmsPageTypePage = IsPage("cmspagetype")
userPage = IsPage("user")
groupPage = IsPage("group")
cmsPagePage = Eval(IsPage("cmspage-report.asp") Or IsPage("cmspage.asp"))

'----------------------------------------------------------------------------------------
' GET GROUP AND PERMISSION LEVELS
'----------------------------------------------------------------------------------------
Dim IS_DEVELOPER, IS_MASTER, IS_MARKET_OWNER, IS_CONTENT_EDITOR, IS_BUSINESS_OWNER
Dim HAS_READ_PERMISSION, HAS_INSERT_PERMISSION, HAS_UPDATE_PERMISSION, HAS_DELETE_PERMISSION, HAS_APPROVAL_PERMISSION

IS_DEVELOPER = IS_MASTER = IS_MARKET_OWNER = IS_CONTENT_EDITOR = IS_BUSINESS_OWNER = False
HAS_READ_PERMISSION = HAS_INSERT_PERMISSION = HAS_UPDATE_PERMISSION = HAS_DELETE_PERMISSION = HAS_APPROVAL_PERMISSION = ""

'----------------------------------------------------------------------------------------
' SET PERMISSION LEVELS
'----------------------------------------------------------------------------------------
Function GetPagePermissionByUrl(url,rp,ip,up,dp,ap)
	Set oCMSPagePerm = New cCMSPagePermission
	Set collCMSPagePerm = New cCMSPagePermission
	collCMSPagePerm.EmailAddress = GetSessionVariable("USER_EMAIL")
	collCMSPagePerm.UrlPage = url
	Call collCMSPagePerm.GetCMSPagePermissionInfoByUrl()
	For Each oCMSPagePerm In collCMSPagePerm.CMSPagePermissions.Items
		rp = rp & CStr(oCMSPagePerm.ReadPermission)
		ip = ip & CStr(oCMSPagePerm.InsertPermission)
		up = up & CStr(oCMSPagePerm.UpdatePermission)
		dp = dp & CStr(oCMSPagePerm.DeletePermission)
		ap = ap & CStr(oCMSPagePerm.ApprovalPermission)
		Set oCMSPagePerm = Nothing
	Next
	Set collCMSPagePerm = Nothing
	
	'Get True/False value pending on permissions
	rp = HasCurrentPermission(rp)
	ip = HasCurrentPermission(ip)
	up = HasCurrentPermission(up)
	dp = HasCurrentPermission(dp)
	ap = HasCurrentPermission(ap)
End Function

'----------------------------------------------------------------------------------------
' IF ON THE ACCESS DENIED PAGE, THEN SKIP ALL OF THIS...
'----------------------------------------------------------------------------------------		
If denyPage = False Then
	
	'----------------------------------------------------------------------------------------
	' IF WE ARE AT THE LOGIN PAGE, THEN SKIP THIS, BUT SET DEFAULT VALUES
	'----------------------------------------------------------------------------------------		
	If loginPage = False Then

		If Not GetSessionVariable("IS_AUTHENTICATED") = True Then
			PageRedirect(CMS_LOGIN_PAGE)
		Else
			'----------------------------------------------------------------------------------------
			' DETERMINE GROUP BY USER
			'----------------------------------------------------------------------------------------
			IS_DEVELOPER = HasGroupPermission(DEVELOPER_GRP_ID)
			IS_MASTER = HasGroupPermission(MASTER_GRP_ID)
			IS_MARKET_OWNER = HasGroupPermission(MKT_OWNER_ID)
			IS_CONTENT_EDITOR = HasGroupPermission(EDITOR_GRP_ID)
			IS_BUSINESS_OWNER = HasGroupPermission(BUS_OWNER_GRP_ID)
			
			If HasPermissionWhileCMSIsLocked = False And CBool(CMS_LOCKED) Then
				echo("<SCR" & "IPT>alert('The CMS has been locked at this time; therefore, you will be redirected to the login page and lose all information you have not saved.\n\nIf you have questions, please contact your CMS system administrator.');</SCR" & "IPT>")
				Call PageRedirect("logout.asp")
			Else
	
				'----------------------------------------------------------------------------------------
				' GET PAGE INFORMATION BASED ON URL ADDRESS
				'
				' :: 	Will get multiple rows back if user belongs to multiple groups; therefore, we
				'		will append the values to each other then look for the "1" bit value in the 
				'		string.
				'----------------------------------------------------------------------------------------		
				Call GetPagePermissionByUrl(PAGE_URL_FILE, _
											HAS_READ_PERMISSION, _
											HAS_INSERT_PERMISSION, _
											HAS_UPDATE_PERMISSION, _
											HAS_DELETE_PERMISSION, _
											HAS_APPROVAL_PERMISSION)
			End If
		End If
		
		'----------------------------------------------------------------------------------------
		' IF NEITHER PAGE IS THE CURRENT PAGE
		'----------------------------------------------------------------------------------------		
		If ( defaultPage = profilePage ) Then	

			'----------------------------------------------------------------------------------------
			' AS WE PROCESS OUR PAGE, WE WILL BE DETERMINING WHETHER THE USER IS DENIED ACCESS
			' TO THE CURRENT PAGE
			'----------------------------------------------------------------------------------------

			'----------------------------------------------------------------------------------------
			' GENERAL PERMISSIONS CHECKPOINT :: MUST HAVE READ PERMISSIONS TO GAIN ACCESS TO ANYTHING
			'----------------------------------------------------------------------------------------		
			If HAS_READ_PERMISSION Then	

				'----------------------------------------------------------------------------------------
				' IF DO NOT HAVE INSERT/UPDATE PERMISSIONS AND ON THE EDIT PAGE
				'----------------------------------------------------------------------------------------					
				If (HAS_INSERT_PERMISSION Or HAS_UPDATE_PERMISSION) = False And IS_EDIT_PAGE Then 
					Call PageRedirect(CMS_ACCESS_DENIED_ERROR_PAGE)				

				'----------------------------------------------------------------------------------------
				' IF HAVE INSERT/UPDATE PERMISSION AND ON THE EDIT PAGE, THEN CHECK GROUP USER IS IN
				'----------------------------------------------------------------------------------------					
				ElseIf (HAS_INSERT_PERMISSION Or HAS_UPDATE_PERMISSION) = True And IS_EDIT_PAGE Then
				
					'----------------------------------------------------------------------------------------
					' IF NOT A DEVELOPER THEN ONLY GIVE ACCESS TO PARTICULAR PAGES IF EDITING AN ENTRY
					'----------------------------------------------------------------------------------------					
					If IS_DEVELOPER = False Then
	
						If StringEmptyOrNull(GetQryString("id")) Then					
							Select Case True
								Case themePage, configPage, cmsPagePage, groupPage
									Call PageRedirect(CMS_ACCESS_DENIED_ERROR_PAGE)																			
							End Select										
						End If
						
					End If

				'----------------------------------------------------------------------------------------
				' IF REPORT PAGE
				'----------------------------------------------------------------------------------------					
				Else 
					Select Case True
						Case themePage, configPage, userPage, cmsPageTypePage, cmsPagePage
							If ( IS_MASTER = False ) Then
								If ( IS_DEVELOPER = False ) Then
									Call PageRedirect(CMS_ACCESS_DENIED_ERROR_PAGE)
								End If
							End If											
					End Select

				End If

			End If						
		End If	
	
	Else
		HAS_READ_PERMISSION = True
		IS_EDIT_PAGE = True
		IS_REPORT_PAGE = True
		Call AddSessionVariable("SITE_LOGIN",1)
	End If

End If

%>
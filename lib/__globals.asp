<%
'----------------------------------------------------------------------------------------
'	WEBSITE DOMAIN: 	http://www.juliestarkphotography.com
'	WEBSITE CONTACT: 	Julie Stark
'	WEBSITE DEVELOPER: 	Brian Gaines ~ bgaines@newleaftechinc.com (800) 670-1487
'	CODE COPYRIGHT:		The code developed for Julie Stark Photography by New Leaf 
'                       Technologies, Inc. is copyrighted and under a general contract
'						agreement. The contract states that if any changes are made to the 
'						codebase by anyone other than Brian Gaines, then the code will 
'						no longer be supported; thus terminating the agreement. Usage 
'						of any code in an application not intended for use on the 
'						www.juliestarkphotography.com website domain is not permissable.
'	CODE VERSION: 		1.0
'----------------------------------------------------------------------------------------

'On Error Resume Next

'----------------------------------------------------------------------------------------
' PREVENT PAGE FROM BEING CACHED BY BROWSER
'----------------------------------------------------------------------------------------
Call NoCache()
SESSION.Timeout = 1000

'----------------------------------------------------------------------------------------
' DEVELOPER INFORMATION
'----------------------------------------------------------------------------------------
Dim DEVELOPER_EMAIL
Dim DEVELOPER_COMPANY
Dim DEVELOPER_WEBSITE
Dim DEVELOPER_PHONE
Dim DEVELOPER_COPYRIGHT
DEVELOPER_EMAIL = "bgaines@newleaftechinc.com"
DEVELOPER_COMPANY = "New Leaf Technologies, Inc."
DEVELOPER_WEBSITE = "http://www.newleaftechinc.com"
DEVELOPER_PHONE = "(800) 670-1487"
DEVELOPER_COPYRIGHT = "The code developed for Julie Stark Photography by New Leaf Technologies, Inc. is copyrighted and under a general contract agreement. The contract states that if any changes are made to the codebase by anyone other than Brian Gaines, then the code will no longer be supported; thus terminating the agreement. Usage of any code in an application not intended for use on the www.juliestarkphotography.com website domain is not permissable."

'----------------------------------------------------------------------------------------
' VARIABLE PREFIXES
'----------------------------------------------------------------------------------------
Dim strCookiePrefix,strCachePrefix
strCachePrefix = "STARK"
strCookiePrefix = strCachePrefix & "_COOK_"			'STARK_COOK_<<COOKIE_NAME>>
strSessionPrefix = strCachePrefix & "_SESS_"		'STARK_SESS_<<SESSION_NAME>>

Sub PageRedirect(path)
	Response.Redirect(path)
End Sub

' Determines if a string is empty or null
Function StringNotEmptyOrNull(strVal)
	If Trim(strVal) & "" <> "" Then
		StringNotEmptyOrNull = True
	Else
		StringNotEmptyOrNull = False
	End If
End Function

' Retrieve a session (user/cached) variable value based on key
Function GetSessionVariable(key)
	if StringNotEmptyOrNull(key) then
		If IsObject(Session(strSessionPrefix & Trim(key))) Then
			Set GetSessionVariable = Session(strSessionPrefix & Trim(key))
		Else
			GetSessionVariable = Session(strSessionPrefix & Trim(key))
		End If
	else
		GetSessionVariable = ""
	end if
End Function

' Create a session (user/cached) variable
Function AddSessionVariable(key,ByVal val)
	On Error Resume Next
	if StringNotEmptyOrNull(key) and StringNotEmptyOrNull(val) then
		If IsObject(val) Then
			Set Session(strSessionPrefix & Trim(key)) = val
		Else
			Session(strSessionPrefix & Trim(key)) = val
		End if
	End If
End Function

' Create an application (global/cached) variable based on key/value pair
Function AddAppVariable(key,ByVal val)
	On Error Resume Next
	if StringNotEmptyOrNull(key) and StringNotEmptyOrNull(val) then
		Application.Lock()
		Application(GetSessionVariable("SITE_LOGIN") & "_" & Trim(UCASE(key))) = Trim(val)
		Application.UnLock()
	End If
End Function

' Retrieve an application (global/cached) variable based on its key
Function GetAppVariable(key)
	if StringNotEmptyOrNull(key) then
		GetAppVariable = Application(GetSessionVariable("SITE_LOGIN") & "_" & Trim(UCase(key)))
		If GetAppVariable & "" <> "" Then
			GetAppVariable = replace(GetAppVariable,"''","'")
		Else
			GetAppVariable = ""
		End If
	else
		GetAppVariable = ""
	end if
End Function

'----------------------------------------------------------------------------------------
' DATABASE CONNECTION STRING << Determine connection -- working locally vs. on server >>
'----------------------------------------------------------------------------------------
Dim CONNECTION_STRING, CONFIG_CONNECTION_STRING
If Request.ServerVariables("SERVER_NAME") = "localhost" Then
	Call AddSessionVariable("SITE_LOGIN","JS")
	CONNECTION_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\Websites\JulieStarkPhotography_PROD\db\jsdb.mdb;"
Else
	Call AddSessionVariable("SITE_LOGIN","JS")
	CONNECTION_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\WebsiteAccessDB\juliestark\jsdb.mdb;"	
End If
%>
<!-- #include virtual="/objects/i_helper.asp" -->
<!-- #include virtual="/objects/cConfig.asp" -->
<%
'Get URL from address bar
Dim url
Dim urlFileName

url = LCase(Trim(Request.ServerVariables("URL")))
PAGE_URL_FILE = Mid(url,InstrRev(url,"/")+1,InstrRev(url,"."))

'----------------------------------------------------------------------------------
'DISCLAIMER NOTICE:
'I took the easy way out and just hard coding the page ids here. I could have 
'performed a page lookup and joined with the meta tag table to get the necessary
'results, but I figured I will change it manually when and IF new pages are added.
'
'PageId	WebPageName
'1		Splash Page
'2		Home Page
'3		About Me
'4		Session Details
'5		Contact Me
'6		Login
'7		Gallery
'8      Proofs (not set up in database)
'----------------------------------------------------------------------------------
Dim WEB_PAGE_ID
WEB_PAGE_ID = 0
Select Case True
	Case Cbool(Instr(PAGE_URL_FILE,"splash")):		WEB_PAGE_ID = 1
	Case Cbool(Instr(PAGE_URL_FILE,"home")): 		WEB_PAGE_ID = 2
	Case Cbool(Instr(PAGE_URL_FILE,"about")): 		WEB_PAGE_ID = 3
	Case Cbool(Instr(PAGE_URL_FILE,"details")): 	WEB_PAGE_ID = 4
	Case Cbool(Instr(PAGE_URL_FILE,"contact")): 	WEB_PAGE_ID = 5
	Case Cbool(Instr(PAGE_URL_FILE,"login")): 		WEB_PAGE_ID = 6
	Case Cbool(Instr(PAGE_URL_FILE,"gallery")): 	WEB_PAGE_ID = 7					
	Case Cbool(Instr(PAGE_URL_FILE,"proofs")):      WEB_PAGE_ID = 8
End Select

Dim WEB_PAGE_NAME
Select Case GetQryString("fid")
	Case "0","":	WEB_PAGE_NAME = "All Pages"
	Case "2": WEB_PAGE_NAME = "Home Page"
	Case "3": WEB_PAGE_NAME = "About Me"
	Case "4": WEB_PAGE_NAME = "Session Details"
	Case "5": WEB_PAGE_NAME = "Contact Me"
	Case "6": WEB_PAGE_NAME = "Login"	
	Case Else: WEB_PAGE_NAME = ""
End Select

'----------------------------------------------------------------------------------------
' GET APPLICATION-LEVEL VARIABLES AND STORE IN GLOBAL CACHE
'----------------------------------------------------------------------------------------
Dim oConfig
Set oConfig = New cConfig
Call oConfig.GetConfigs()
Set oConfig = Nothing

Dim ROOT_PATH
Dim IMAGE_PATH
Dim DOMAIN_NAME
Dim MAIL_OBJECT
Dim MAIL_SERVER_IP_ADDRESS
Dim GALLERY_PATH
Dim PROOF_PATH
Dim COMPANY_NAME
Dim COMPANY_ADDRESS
Dim COMPANY_CITY
Dim COMPANY_STATE
Dim COMPANY_ZIP
Dim COMPANY_PHONE
Dim COMPANY_FAX
Dim COMPANY_EMAIL
Dim COMPANY_LOGO
Dim PHOTOGRAPHER_FNAME
Dim PHOTOGRAPHER_LNAME
Dim PHYSICAL_ROOT_PATH
Dim FEXT_ALLOWED
Dim MAX_FILE_KB_UPLOAD_SIZE
Dim DISABLE_IMAGE_RIGHT_MOUSE_CLICK
Dim MAX_PX_GALLERY_IMAGE_HEIGHT
Dim MAX_PX_GALLERY_IMAGE_WIDTH
Dim MAX_PX_THUMBNAIL_HEIGHT
Dim MAX_PX_THUMBNAIL_WIDTH
Dim GALLERY_TRANS_FACTOR
Dim FILE_UPLOAD_BATCH_COUNT
Dim SCRIPT_TIMEOUT_IN_MINUTES
Dim SLIDE_SHOW_SLIDE_SPEED
Dim SLIDE_SHOW_DURATION_RATE
Dim SLIDE_SHOW_MOTION
Dim SLIDE_SHOW_GRADIENT_SIZE
Dim RUN_SLIDE_SHOW
Dim RANDOMIZE_SIDE_PHOTO
Dim SLIDE_SHOW_SIDE_PHOTO
Dim SLIDE_SHOW_WIPE_STYLE
Dim DAYS_FROM_NOW_TO_EXPIRE_ACCT

PHYSICAL_ROOT_PATH = GetAppVariable("PHYSICAL_ROOT_PATH")
ROOT_PATH =	GetAppVariable("ROOT_PATH")
IMAGE_PATH = GetAppVariable("IMAGE_PATH")
DOMAIN_NAME = GetAppVariable("DOMAIN_NAME")
MAIL_OBJECT = GetAppVariable("MAIL_OBJECT")
MAIL_SERVER_IP_ADDRESS = GetAppVariable("MAIL_SERVER_IP_ADDRESS")
GALLERY_PATH = GetAppVariable("GALLERY_PATH")
PROOF_PATH = GetAppVariable("PROOF_PATH")
COMPANY_NAME = GetAppVariable("COMPANY_NAME")
COMPANY_ADDRESS	= GetAppVariable("COMPANY_ADDRESS")
COMPANY_CITY = GetAppVariable("COMPANY_CITY")
COMPANY_STATE = GetAppVariable("COMPANY_STATE")
COMPANY_ZIP	= GetAppVariable("COMPANY_ZIP")
COMPANY_PHONE = GetAppVariable("COMPANY_PHONE")
COMPANY_FAX	= GetAppVariable("COMPANY_FAX")
COMPANY_EMAIL = GetAppVariable("COMPANY_EMAIL")
COMPANY_LOGO = GetAppVariable("COMPANY_LOGO")
PHOTOGRAPHER_FNAME = GetAppVariable("PHOTOGRAPHER_FNAME")
PHOTOGRAPHER_LNAME = GetAppVariable("PHOTOGRAPHER_LNAME")
MAX_FILE_KB_UPLOAD_SIZE = GetAppVariable("MAX_FILE_KB_UPLOAD_SIZE")
FEXT_ALLOWED = GetAppVariable("FEXT_ALLOWED")
DISABLE_IMAGE_RIGHT_MOUSE_CLICK = GetAppVariable("DISABLE_IMAGE_RIGHT_MOUSE_CLICK")
MAX_PX_GALLERY_IMAGE_HEIGHT = GetAppVariable("MAX_PX_GALLERY_IMAGE_HEIGHT")
MAX_PX_GALLERY_IMAGE_WIDTH = GetAppVariable("MAX_PX_GALLERY_IMAGE_WIDTH")
MAX_PX_THUMBNAIL_HEIGHT = GetAppVariable("MAX_PX_GALLERY_IMAGE_HEIGHT")
MAX_PX_THUMBNAIL_WIDTH = GetAppVariable("MAX_PX_GALLERY_IMAGE_WIDTH")
GALLERY_TRANS_FACTOR = GetAppVariable("GALLERY_TRANS_FACTOR")
FILE_UPLOAD_BATCH_COUNT = GetAppVariable("FILE_UPLOAD_BATCH_COUNT")
SCRIPT_TIMEOUT_IN_MINUTES = GetAppVariable("SCRIPT_TIMEOUT_IN_MINUTES")
SLIDE_SHOW_SLIDE_SPEED = GetAppVariable("SLIDE_SHOW_SLIDE_SPEED")
SLIDE_SHOW_DURATION = GetAppVariable("SLIDE_SHOW_DURATION")
SLIDE_SHOW_MOTION = GetAppVariable("SLIDE_SHOW_MOTION")
SLIDE_SHOW_GRADIENT_SIZE = GetAppVariable("SLIDE_SHOW_GRADIENT_SIZE")
RUN_HOME_PAGE_SLIDE_SHOW = GetAppVariable("RUN_HOME_PAGE_SLIDE_SHOW")
RANDOMIZE_SIDE_PHOTO = GetAppVariable("RANDOMIZE_SIDE_PHOTO")
RUN_SIDE_PHOTO_SLIDE_SHOW = GetAppVariable("RUN_SIDE_PHOTO_SLIDE_SHOW")
SLIDE_SHOW_WIPE_STYLE = GetAppVariable("SLIDE_SHOW_WIPE_STYLE")
DAYS_FROM_NOW_TO_EXPIRE_ACCT = GetAppVariable("DAYS_FROM_NOW_TO_EXPIRE_ACCT")

Dim ALT_IMAGE_TEXT
ALT_IMAGE_TEXT = COMPANY_NAME & ", " & COMPANY_CITY & " " & COMPANY_STATE

'SERVER/PATH INFORMATION
If Request.ServerVariables("SERVER_NAME") = "localhost" Then
	ROOT_PATH = "/"
	PHYSICAL_ROOT_PATH = PHYSICAL_ROOT_PATH & "\JulieStarkPhotography"
Else
	ROOT_PATH = DOMAIN_NAME & GetAppVariable("ROOT_PATH")
	PHYSICAL_ROOT_PATH = PHYSICAL_ROOT_PATH & Replace(ROOT_PATH,"/","\")
End If
IMAGE_PATH = ROOT_PATH & GetAppVariable("IMAGE_PATH")
%>

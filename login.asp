<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/objects/cLogin.asp" -->
<!-- #include virtual="/objects/cPhotos.asp" -->
<!-- #include virtual="/objects/cGallery.asp" -->
<%
'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim oLogin
Dim strUser
Dim strPwd
Dim errmsg

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function GetUserByLogin(login,pwd)
	Dim oLogin
	Dim collLogin
	Set oLogin = New cLogin
	Set collLogin = New cLogin
	With collLogin
		.Login = login
		.Password = pwd
		Call .GetUserByLogin()
		For Each oLogin In .Logins.Items
			If oLogin.ID > 0 Then
				If Now() > oLogin.Expire Then
					errmsg = "Your account login has expired. Please contact " & oLogin.Photographer & " if you need to regain access to your account."
				Else
					If oLogin.IsAdmin Then
						Call AddSessionVariable("IS_ADMIN_USER",oLogin.IsAdmin)
						Call AddSessionVariable("FULLNAME",oLogin.FullName)				    					
						PageRedirect("/cms/gallery-report.asp")
					Else					
						If StringNotEmptyOrNull(oLogin.GalleryName) Then						
							Dim lockMessage
							Dim blnLocked
							Dim blnInactiveGallery
                            blnInactiveGallery = False
							Set oGallery = New cGallery
							Set collGallery = New cGallery							
							collGallery.GalleryLastName = GetSessionVariable("PROOF_GALLERY_NAME")
							Call collGallery.GetGalleryByLastName()
							For Each oGallery In collGallery.Galleries.Items		
								If Now() > oGallery.ExpirationDate Or oGallery.ActiveFlag = False Then
									Set oGallery = Nothing
									blnLocked = True
                                    blnInactiveGallery = (oGallery.ActiveFlag)
									Call AddSessionVariable("PROOF_GALLERY_LOCKED","YES")
									Exit For
								End If
								Set oGallery = Nothing
							Next
							Set collGallery = Nothing
									
							If blnLocked Then
								errmsg = "<span>Your proofing gallery has expired.  Please contact me at " & COMPANY_PHONE & " if you are interested in having your gallery republished.</span><br>"
						    ElseIf blnInactiveGallery = True Then
						        errmsg = "<span>Your proofs are inactive.  Please contact me at " & COMPANY_PHONE & " if you require them to be activated again.</span><br>"
							Else		
								Call AddSessionVariable("PROOF_GALLERY_LOCKED","NO")
								Call AddSessionVariable("PROOF_GALLERY_NAME",LCase(oLogin.GalleryName))
							    PageRedirect("proofs.asp")
							End If
						Else
							errmsg = "<span>An online proofing gallery was not found under your account.</span><br>"
						End If
					End If
				End If
			Else
				errmsg = "<span>You entered an incorrect username and/or password. please try again</span><br>"
				Set oLogin = Nothing
				Exit For
			End If
			Set oLogin = Nothing
		Next
	End With
	Set collLogin = Nothing
End Function


'----------------------------------------------------------------------------------------
' AUTHENTICATION PROCESS
'----------------------------------------------------------------------------------------
If GetFormPost("Login") = "Log In" Then
	strUser = GetFormPost("txtUserName")
	strPwd = GetFormPost("txtPassword")
	Call GetUserByLogin(strUser,strPwd)	
End If
 
Dim BODY_TITLE
BODY_TITLE = "Client Login"
%>
<html>
<head>
<title>Log into the <%=COMPANY_NAME%> online proofing gallery to view your online proofs.</title>
<!-- #include file="__meta.asp" -->
<!-- #include file="__css.asp" -->
<script src="__menu.js" type="text/javascript"></script>
</head>
<body onload="preloadImages();">
<form method="post" action="login.asp" name="form1">
<table width="900" height="562" border="0" cellpadding="0" cellspacing="0" align="center">
    <!-- #include file="__menu.asp" -->
        <td width="826" height="359" colspan="11" valign="top" align="center" style="background-color:<%=BodyContentBackColor%>;">
            <table width="826">
		        <tr>
		            <td><img src="images/spacer.gif" width="1" height="340" alt="" /></td>
		            <td class="body-content" valign="top" width="510">
			            <table cellspacing="0" border="0" class="body-contact-form" width="530">
				            <tr>
				                <td><img src="images/spacer.gif" width="1" height="316" alt="" /></td>
					            <td align="center" width="575">
						            <h1><%=BODY_TITLE%></h1>
						            <p style="padding:0 5 0 0;" align="left">Please enter your user name and password to access your online proof gallery. If you have not yet received this information, please <a href="contact.asp" style="color:#554120">contact <%=PHOTOGRAPHER_FNAME%></a> using the online web form.</p>
				                    <% echo(errmsg) %>
				                    <table width="400" border="0" align="center" cellspacing="0" cellpadding="3">
					                    <tr>
						                    <td width="100">User Name</td>
						                    <td width="250"><input type="text" name="txtUserName" style="background-color:#F5F5DD;width:150px;font:10px #666 tahoma" value=""></td>
					                    </tr>
					                    <tr>
						                    <td width="100">Password</td>
						                    <td width="250"><input type="password" name="txtPassword" style="background-color:#F5F5DD;width:150px;font:10px #666 tahoma" value=""></td>
					                    </tr>
					                    <tr>
					                        <td>&nbsp;</td>
						                    <td align="left"><input type="submit" style="color:#fff;font-size:10px;font-weight:bold;border-style:solid;border-width:1px;border-color:#554120;padding:2 10 2 10;background-color:#554120;" name="Login" value="Log In"></td>
					                    </tr>
					                </table>
    				            </td>
				            </tr>
			            </table>		
                    </td>			                    
                    <!--#include file="__sidephoto.asp"-->
                </tr>
		    </table>	
        </td>
<!-- #include file="__partrow8_row9_10_copyright.asp" -->
</form>
</body>
</html>
<script type="text/javascript" language="javascript"><!--
document.form1.txtUserName.focus();
//--></script>
<% 
If RUN_HOME_PAGE_SLIDE_SHOW And WEB_PAGE_ID = 2 Then %>
<script type="text/javascript" language="javascript"><!--	
var ie = document.all;
var whichimage=0;
var blenddelay=(ie) ? document.images.slide.filters[0].duration*1000 : 0;
function slideit() {
	if (!document.images) return
	if (ie) document.images.slide.filters[0].apply();
	document.images.slide.src = imageholder[whichimage].src;
	if (ie) document.images.slide.filters[0].play();
	whichimage =(whichimage<slideimages.length-1) ? whichimage+1 : 0;
	setTimeout("slideit()",slidespeed+500);
}
slideit();
//--></script>
<% End If %>
<!-- #include file="__disable_mouse_click.asp" -->

<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/objects/cPhotos.asp" -->
<html>
<head>
<title>contact julie stark children photography in New York to learn more about the services we offer</title>
<title>contact julie stark children photography in New York to learn more about the services we offer</title>
<!-- #include file="__meta.asp" -->
<!-- #include file="__css.asp" -->
<% 

Function SendSAMail(subj,body,fromName,toName)
	'Dim mailer
	'Set mailer = Server.CreateObject("SoftArtisans.SMTPMail")
	'mailer.RemoteHost = "216.82.113.58"
	'mailer.FromAddress = "julie@juliestarkphotography.com"
	'mailer.AddRecipient "", "briannlt@gmail.com"
	'mailer.Subject = "subject" & " - " & FormatDateTime(Now(),1)
	'mailer.HtmlText = "test body"
	'mailer.SendMail
	'set mailer = nothing

        Dim mailer
        Set mailer = CreateObject("SoftArtisans.SMTPMail")
        With mailer
            .RemoteHost = "mail.juliestarkphotography.com"
            .Subject = Now()
            .HtmlText = "test"
            .AddRecipient "", "briannlt@gmail.com"
            .FromAddress = "website@juliestarkphotography.com"
             mailer.SendMail()
        End With
End Function

Function SendMail(subj,body,fromName,toName)   
	Dim objMsg
	Const SMTP_SERVER = "216.82.113.58"
	Set objMsg = Server.CreateObject("CDO.Message")
	With objMsg
		.To = toName
		.From = fromName
		.Subject = subj & " - " & FormatDateTime(Now(),1)
		.HTMLBody  = body
		On Error Resume Next
		.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")= 2
		.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")= SMTP_SERVER
		.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25 
		.Configuration.Fields.Update
		.Send
	End With

    Set objMsg = Nothing
    If err <> 0 Then
        ' error - this is our custom error message
        ' errNum - the error number 
        ' errDesc - the error description
        SendMail = False
    Else
        SendMail = True
    End If	
End Function

Dim BODY_TITLE
BODY_TITLE = "Contact " & PHOTOGRAPHER_FNAME  

' check to see if we have answers to our survey
If Len(Request.Form("Full Name")) > 0 Then
	Dim strWebForm, ix, strCss, strThankYou

	' loop through each form element in order as it appears to the user
	For ix = 1 to Request.Form.Count
	
		' get the name and value of each form
		fieldName = Request.Form.Key(ix)
		fieldValue = Request.Form.Item(ix) 

		' we don't want to see the submit button in the email
		if fieldName <> "x" and fieldName <> "y" then

			if fieldValue = "" then fieldValue = "(empty)"
			strWebForm = strWebForm & fieldName & ": " & fieldValue & "<br>"
			
		End If
		
		' move to the next form element in the collection
	Next	
	
	strCss = "<style>"
	strCss = strCss & "body,p,h1,h2{font-family:Verdana, Arial, Sans-Serif; font-size:12px;}"
	strCss = strCss & "h1{font-size:16px; font-weight:bold;}"
	strCss = strCss & "h2{font-size:14px; font-weight:bold;}"
	strCss = strCss & "</style>"
	
	' final preparations to the email message
	strWebForm = "<html><head>" & strCss & "</head><body>" & PHOTOGRAPHER_FNAME & ",<br><br>The following information was entered from your website. Please review the comments and reply to the sender if necessary:<br><br><p>" & strWebForm & "</p></body></html>"

        Dim mailer
        Set mailer = CreateObject("SoftArtisans.SMTPMail")
        With mailer
            	.RemoteHost = "mail.juliestarkphotography.com"
            	.Subject = "Email from JulieStarkPhotography.com - " & FormatDateTime(Now(),1)
            	.HtmlText = strWebForm
            	.AddRecipient "", "julie@juliestarkphotography.com"
            	.FromAddress = "website@juliestarkphotography.com"
		.ReplyTo  = "julie@juliestarkphotography.com"
		.UserName= "website@juliestarkphotography.com"
		.Password= "website"
             	.SendMail()
        End With

	strThankYou = "<p>Thank you, your message has been sent.</p>"	

	'Call SendSAMail("Email from JulieStarkPhotography.com - " & FormatDateTime(Now(),1),PHOTOGRAPHER_FNAME & ",<br><br>The following information was entered from your website:<br><br>" & strWebForm,"julie@juliestarkphotography.com","bgaines@newleaftechinc.com")
	'Response.end

	'If SendSAMail("Email from JulieStarkPhotography.com - " & FormatDateTime(Now(),1),PHOTOGRAPHER_FNAME & ",<br><br>The following information was entered from your website:<br><br>" & strWebForm,"julie@juliestarkphotography.com","bgaines@newleaftechinc.com") = False Then
	    'strThankYou = "<p>Thank you, your message has been sent.</p>"
	'Else
	    'Response.Write(err.Description)
	'End If
End If
%>
<script type="text/javascript" language="javascript">
<!--
function send() {
	document.form1.submit();
}
//-->
</script>
<script src="__menu.js" type="text/javascript"></script>
</head>
<body onload="preloadImages();">
<table width="900" height="562" border="0" cellpadding="0" cellspacing="0" align="center"> 
    <!-- #include file="__menu.asp" -->
        <td width="826" height="359" colspan="11" valign="top" align="center" style="background-color:<%=BodyContentBackColor%>;">
            <table width="826">
		        <tr>
		            <td><img src="images/spacer.gif" width="1" height="340" alt="" /></td>
                    <!--#include file="__sidephoto.asp"-->
		            <td class="body-content" valign="top" width="510">
			            <table cellspacing="0" border="0" class="body-contact-form" width="530">
				            <tr>
				                <td><img src="images/spacer.gif" width="1" height="316" alt="" /></td>
					            <td align="center" width="575">
						            <h1><%=BODY_TITLE%></h1>
						            <p style="padding:0 5 0 0;" align="left">Please complete the form below to request additional information about <%=PHOTOGRAPHER_FNAME%>'s services or to request a user name and password for your proofing gallery.<%=strThankYou%></p>
					                <form action="contact.asp?send=1" method="post" name="form1">
					                    <table width="400" border="0" align="center" cellspacing="0" cellpadding="3">
						                    <tr>
							                    <td width="100">Full Name</td>
							                    <td width="250"><input type="text" name="Full Name" style="background-color:#F5F5DD;width:150px;font:10px #666 tahoma" value=""></td>
						                    </tr>
						                    <tr>
							                    <td width="100">Email</td>
							                    <td width="250"><input type="text" name="Email Address" style="background-color:#F5F5DD;width:150px;font:10px #666 tahoma" value=""></td>
						                    </tr>
						                    <tr>
							                    <td width="100">Telephone</td>
							                    <td width="250"><input type="text" name="Phone Number" style="background-color:#F5F5DD;width:150px;font:10px #666 tahoma" value=""></td>
						                    </tr>
						                    <tr>
							                    <td width="100">Message</td>
							                    <td width="250"><textarea name="Comments" cols="40" rows="5" style="background-color:#F5F5DD;font:12px #666 tahoma" wrap="virtual"></textarea></td>
						                    </tr>
						                    <tr>
						                        <td>&nbsp;</td>
							                    <td align="left"><a style="color:#fff;font-size:10px;font-weight:bold;border-style:solid;border-width:1px;border-color:#554120;padding:2 10 2 10;background-color:#554120;" href="javascript:send();">Request Information</a></td>
						                    </tr>
						                </table>
					                </form>
					            </td>
				            </tr>
			            </table>		
                    </td>
		        </tr>
		    </table>	
        </td>
<!-- #include file="__partrow8_row9_10_copyright.asp" -->
</body>
</html>

<% 
If RUN_HOME_PAGE_SLIDE_SHOW And WEB_PAGE_ID = 2 Then %>
<script language="JavaScript1.1"><!--	
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

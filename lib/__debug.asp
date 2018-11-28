<%
Dim intBorder
Dim strHeadColor
Dim strLeftNavColor
Dim strBodyColor
Dim strFooterColor
Dim color
Dim strAdminHeaderColor
Dim strTableHeader
Dim strTopBarColor
Dim strLeftNavHeaderClassName
Dim strTopDividerBarImage
Dim strTableHeaderFontColor 

intBorder = 0 ' set this value to 0 when you want to remove the table borders

strHeadColor = "#FFFFFF"
'	strLeftNavColor = "#8DBCEB"
strAdminHeaderColor = "#99B27F"
strLeftNavColor = "beige"
strBodyColor = "#FFFFFF"
strFooterColor = "#EEEEEE"
strTopBarColor = "#8DBCEB"
strTableHeader = "#8DBCEB"
strTableHeaderFontColor = "fontBlack"
strLeftNavTableColor = "#FF8800"
strMouseOverColor ="#CCCCCC"
strLeftNavHeaderClassName = "LeftNavHeaderBlack"
'strTopDividerBarImage = GetAppVariable(APPVARNAME & "AppTopLevelImageUploadPath") & "/site/gradient_dkblue_left.gif"
strTopDividerBarImage = strImagePath & "/site/gradient_dkblue_left.gif"
	
strTableHeader = "#4A6999"
strTableHeaderFontColor = "fontWhite"
strMouseOverColor ="#FFCC66"

' PURPOSE:	Iterate through cookie collection and write results
Function GetCookieJar()
	On Error Resume Next
	For Each objItem In request.Cookies
		If request.Cookies(objItem).HasKeys Then
			' loop through the subkeys of the cookie collection
			For Each objItemKey in request.Cookies(objItem)
				response.Write objItem & "(" & objItemKey & ") = " & request.Cookies(objItem)(objItemKey) & "<br/>"
			Next
		else
			' print out the cookie string as normal
			response.write objItem & " = " & request.Cookies(objItem) & "<br/>"
		end if
	Next
End Function

' PURPOSE:	Iterate through QueryString collection and write results
Function GetQryStringCollection()
	On Error Resume Next
	For Each objItem In request.QueryString
		' print out the querystring key/value pairs
		response.write objItem & " = " & request.QueryString(objItem) & "<br/>"
	Next
End Function

' PURPOSE:	Iterate through Forms collection and write results
Function GetFormPostCollection()
	On Error Resume Next
	For Each objItem In request.Form
		' print out the form key/value pairs
		response.write objItem & " = " & request.QueryString(objItem) & "<br/>"
	Next
End Function

Function SendErrorEmail(ErrorType, ErrorSource, ErrorNumber, ErrorDescription)

	On Error Resume Next

   	'Declare variables
    Dim HTML,CollectionItem,iNumber,myMail,QS,RF

    'Transfer the contents of the QueryString and Form collections to variables
    Set QS = Request.QueryString
    Set RF = Request.Form

    'Declare constants.
    Const MAIL_FROM_EMAIL = "gainesme72@attbi.com" 'Email address of email sender
    Const MAIL_TO_EMAIL = "gainesme72@attbi.com" 'Email address of email recipient
    Const MAIL_SUBJECT = "Error Report" 'Title of error report

    'Generate the top part of the error report
    HTML = "<!DOCTYPE HTML PUBLIC""-//IETF//DTD HTML//EN"">"
    HTML = HTML & "<html>"
    HTML = HTML & "<head>"
    HTML = HTML & "<title>" & MAIL_SUBJECT & "</title>"
    HTML = HTML & "</head>"
    HTML = HTML & "<body bgcolor=""FFFFFF"">"
    HTML = HTML & "<p><font size =""3"" face=""Century Gothic, Arial"">"
    HTML = HTML & "<b>" & MAIL_SUBJECT & "</b><br>"
    HTML = HTML & "Error Report Generated: <FONT COLOR=""#3333FF"">" & FormatDateTime(now(), vbLongDate) & ", " & FormatDateTime(now(), vbLongTime) & "</font><br>"
    HTML = HTML & "<hr>"

    'Generate the error report general description
    HTML = HTML & "<b>Details:</b><br>"
    HTML = HTML & "Error In Page: <FONT COLOR=""#FF3333"">" & Request.ServerVariables("PATH_INFO") & "</FONT><BR>"
    HTML = HTML & "Error Type: <FONT COLOR=""#FF3333"">" & ErrorType & "</FONT><BR>"
    HTML = HTML & "Error Source: <FONT COLOR=""#FF3333"">" & ErrorSource & "</FONT><BR>"
    HTML = HTML & "Error Number: <FONT COLOR=""#FF3333"">" & ErrorNumber & "</FONT><BR>"
    HTML = HTML & "Error Description: <FONT COLOR=""#FF3333"">" & ErrorDescription & "</FONT><BR>"
    HTML = HTML & "<hr>"

    'Report the contents of the QueryString collection
    HTML = HTML & "<b>QueryString Collection:</b><br>"
    If QS.Count > 0 Then
    For Each CollectionItem In QS
            HTML = HTML & CollectionItem & " : <FONT COLOR=""#3333FF"">" & QS(CollectionItem) & "</FONT><br>"
    Next
    Else
    HTML = HTML & "<FONT COLOR=""#FF3333"">The QueryString collection is empty</FONT><br>"
    End If

    HTML = HTML & "<hr>"

   'Report the contents of the Form collection
    HTML = HTML & "<b>Form Collection:</b><br>"

    If RF.Count > 0 Then
		For Each CollectionItem In RF
			HTML = HTML & CollectionItem & " : <FONT COLOR=""#FF3333"">" & RF(CollectionItem) & "</FONT><br>"
		Next
    Else
	    HTML = HTML & "<FONT COLOR=""#3333FF"">The Form collection is empty</FONT><br>"
    End If

    HTML = HTML & "<hr>"

   'Report the Server object properties
    HTML = HTML & "<b>Server Settings:</b><br>"
    HTML = HTML & "ScriptTimeout: <FONT COLOR=""#FF3333"">" & Server.ScriptTimeout & "</FONT><BR>"

    HTML = HTML & "<hr>"

    'Report the Session object properties and the contents of the Session collection
    'IMPORTANT: If you have disabled Sessions either in IIS or
    'by use of the @ENABLESESSIONSTATE = FALSE directive then you MUST comment out this section
    HTML = HTML & "<b>Session Settings:</b><br>"
    HTML = HTML & "CodePage: <FONT COLOR=""#FF3333"">" & Session.CodePage & "</FONT><BR>"
    HTML = HTML & "LCID: <FONT COLOR=""#FF3333"">" & Session.LCID & "</FONT><BR>"
    HTML = HTML & "SessionID: <FONT COLOR=""#FF3333"">" & Session.SessionID & "</FONT><BR>"
    HTML = HTML & "Timeout: <FONT COLOR=""#FF3333"">" & Session.TimeOut & "</FONT><BR>"

    HTML = HTML & "<hr>"

    HTML = HTML & "<b>Session Collection:</b><br>"

    For iNumber = 1 To Session.Contents.Count
        If IsObject(Session.Contents(iNumber)) Then
            HTML = HTML & Session.Contents.Key(iNumber) & "<FONT COLOR=""#3333FF"">[Object]</FONT><BR>"
        Else
            If IsArray(Session.Contents(iNumber)) Then
                HTML = HTML & Session.Contents.Key(iNumber) & "<FONT COLOR=""#3333FF"">[Array]</FONT><BR>"
            Else
                HTML = HTML & Session.Contents.Key(iNumber) & ": <FONT COLOR=""#3333FF"">" & Session.Contents(iNumber) & "</FONT><BR>"
            End If
        End If
    Next

    HTML = HTML & "<hr>"

   'Report the contents of the Application collection
    HTML = HTML & "<b>Application Collection:</b><br>"

    For iNumber = 1 To Application.Contents.Count
        If IsObject(Application.Contents(iNumber)) Then
            HTML = HTML & Application.Contents.Key(iNumber) & "<FONT COLOR=""#3333FF"">[Object]</FONT><BR>"
        Else
            If IsArray(Application.Contents(iNumber)) Then
                HTML = HTML & Application.contents.Key(iNumber) & "<FONT COLOR=""#3333FF"">[Array]</FONT><BR>"
            Else
                HTML = HTML & Application.contents.Key(iNumber) & ": <FONT COLOR=""#3333FF"">" & Application.Contents(iNumber) & "</FONT><BR>"
            End If
        End If
    Next

    HTML = HTML & "<hr>"

   'Report the contents of the Server Variables collection
    HTML = HTML & "<b>Server Variables:</b><br>"

    For Each CollectionItem in request.servervariables
        If CollectionItem <> "ALL_HTTP" and CollectionItem <> "ALL_RAW" then
            HTML = HTML & CollectionItem & " : <FONT COLOR=""#3333FF"">" & request.servervariables(CollectionItem) & "</FONT><br>"
        End If
    Next

    HTML = HTML & "</body>"
    HTML = HTML & "</html>"

	Set objMail = Server.CreateObject("CDONTS.NewMail")

	With objMail
		.From = MAIL_FROM_EMAIL
		.To = MAIL_TO_EMAIL
		.Subject = MAIL_SUBJECT & " at " & Now
		.Body = HTML
		.MailFormat = 0
		.BodyFormat = 0
		.Importance = 2
		.Send
		Set objMail = Nothing
	End With
	
	if err.number <> 0 then
		SendErrorEmail = False
	else
		SendErrorEmail = True
	End If

End Function


Sub StartTimer(x)
'::::::::::::::::::::::::::::::::::::::::::::::::::::
':::     Here we start the timer                  :::
'::::::::::::::::::::::::::::::::::::::::::::::::::::
   StopWatch(x) = timer
end sub


Function StopTimer(x)
'::::::::::::::::::::::::::::::::::::::::::::::::::::
':::     Here we 'stop' the timer and calculate   :::
':::     the elapsed time (allowing for the       :::
':::     midnight wrap-around. NOTE: These        :::
':::     routines should not be used to time      :::
':::     very long events (>1 hour) or very       :::
':::     short events (< 100 milliseconds)        :::
'::::::::::::::::::::::::::::::::::::::::::::::::::::
	Dim EndTime

	EndTime = Timer
	
	'Watch for the midnight wraparound...
	if EndTime < StopWatch(x) then
	 EndTime = EndTime + (86400)
	end if
	
	StopTimer = EndTime - StopWatch(x)
end function

%>
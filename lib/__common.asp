<%


Dim debug
Sub AppendDebugInfo(key,value)

	debug = debug & UCASE(key) & ": " & value & "<BR>"

End Sub

Sub SendDebugInfo
        Dim mailer
        Set mailer = CreateObject("SoftArtisans.SMTPMail")
        With mailer
            	.RemoteHost = "mail.domain.com"
            	.Subject = "Email from domain.com - " & FormatDateTime(Now(),1)
            	.HtmlText = debug
            	.AddRecipient "", "me@domain.com"
            	.FromAddress = "website@domain.com"
		.ReplyTo  = "reply@domain.com"
		.UserName= "website@domain.com"
		.Password= ""
             	.SendMail()
        End With
End Sub


'--------------------------------------------------------------------------------------
' CREATED BY: BRIAN GAINES
' FILENAME: 	__common.asp
' PURPOSE:		This file contains common routines that are used in a web application
'--------------------------------------------------------------------------------------
function SingleQuotes(pStringIn)
    if pStringIn = "" or isnull(pStringIn) then 
	SingleQuotes = Null
	exit function
    end if
	Dim pStringModified
    pStringModified = Replace(pStringIn,"'","''")
    SingleQuotes =  pStringModified
end function

function LineBreakReplace(inString)
	LineBreakReplace = Replace(inString,chr(10) & chr(13),"<br>")
End function

public function echo(p_STR)
    response.write p_Str
end function

public function die(p_STR)
    echo p_Str
    response.end
end function

public function echobr(p_STR)
    echo p_Str & "<br>" & vbCRLF
end function

public function htmlencode(p_STR)
    htmlencode = trim(server.htmlencode(p_Str & " "))
end function

Randomize 'Insure that the numbers are really random
Function RandomString(p_NumChars)
    Dim n
    Dim tmpChar,tmpString
    for n = 0 to p_NumChars
        tmpChar = Chr(Int(32+( Rnd * (126-33))))
        'Random characters (letters, numbers, etc.)
        tmpString = tmpString & tmpChar
    next
    RandomString = tmpString
End Function

Function RenderRequiredFieldsMessageRow()
	RenderRequiredFieldsMessageRow = "<tr><td>&nbsp;</td><td><em>*</em><i>The following fields are required.</i></td></tr>"
End Function

'----------------------------------------------------------------------
' FUNCTION NAME:	CloseCmd
' CREATED BY:		Brian Gaines
' CREATED ON:		11/5/2002
' INPUTS:			cmd - command object
' PURPOSE:			Closes the command object
'----------------------------------------------------------------------
function CloseCmd(cmd)
	if Not cmd.State = adStateOpen Then
		If cmd.State = adStateOpen Then
			Set cmd.ActiveConnection = Nothing
		End If
		Set cmd = Nothing
	End If
end function

'----------------------------------------------------------------------
' FUNCTION NAME:	CloseDBConn
' CREATED BY:		Brian Gaines
' CREATED ON:		11/5/2002
' INPUTS:			conn - connection object
' PURPOSE:			Closes the database connection object, which will
'					free up the connection object to the pool
'----------------------------------------------------------------------
function CloseDBConn(conn)
	If conn.State = adStateOpen Then
		conn.Close
	End If
end function

'----------------------------------------------------------------------
' FUNCTION NAME:	CloseRS
' CREATED BY:		Brian Gaines
' CREATED ON:		11/5/2002
' INPUTS:			recSet
' PURPOSE:			Closes recordset object if open
'----------------------------------------------------------------------
function CloseRS(recSet)
	If Not recSet Is Nothing Then
		If recSet.State = adStateOpen Then
			  recSet.Close
		End	If
		Set recSet = Nothing
	End If
end function

' CREATED BY: 	Originally by some guy named Dan.	
' UPDATED BY:		Ken Schaefer (http://www.4guysfromrolla.com)
' LINK: 				http://www.4guysfromrolla.com/webtech/022701-1.shtml
'
' PURPOSE: 			Gets the ordinal day based on the numeric value that is passed
' INPUTS:				intDay - The numic day of the month.
Function GetDayOrdinal(byVal intDay)
 ' Accepts a day of the month as an integer and returns the appropriate suffix

	Dim strOrd
	
	Select Case intDay
		Case 1, 21, 31
			strOrd = "st"
		Case 2, 22
			strOrd = "nd"
		Case 3, 23
			strOrd = "rd"
		Case Else
			strOrd = "th"
	End Select

	GetDayOrdinal = strOrd

End Function 


' PURPOSE: 	Accepts strDate as a valid date/time, strFormat as the output template.
' 			The function finds each item in the template and replaces it with the
' 			relevant information extracted from strDate
' INPUTS: 	Argument list of occupying any of the following delimiters
'
' 			%m Month as a decimal (02)
' 			%B Full month name (February)
' 			%b Abbreviated month name (Feb )
' 			%d Day of the month (23)
' 			%O Ordinal of day of month (eg st or rd or nd)
' 			%j Day of the year (54)
' 			%Y Year with century (1998)
' 			%y Year without century (98)
' 			%w Weekday as integer (0 is Sunday)
' 			%a Abbreviated day name (Fri)
' 			%A Weekday Name (Friday)
' 			%H Hour in 24 hour format (24)
' 			%h Hour in 12 hour format (12)
' 			%N Minute as an integer (01)
' 			%n Minute as optional if minute <> 0
' 			%S Second as an integer (55)
' 			%P AM/PM Indicator (PM)
Function FormatDate(byVal strDate, byRef strFormat)

	On Error Resume Next
	
	Dim intPosItem, int12HourPart, str24HourPart, strMinutePart, strSecondPart, strAMPM
	
	' Insert Month Numbers
	strFormat = Replace(strFormat, "%m", DatePart("m", strDate), 1, -1, vbBinaryCompare)
	
	' Insert non-Abbreviated Month Names
	strFormat = Replace(strFormat, "%B", MonthName(DatePart("m", strDate), False), 1, -1, vbBinaryCompare)
	
	' Insert Abbreviated Month Names
	strFormat = Replace(strFormat, "%b", MonthName(DatePart("m", strDate), True), 1, -1, vbBinaryCompare)
	
	' Insert Day Of Month
	strFormat = Replace(strFormat, "%d", DatePart("d",strDate), 1, -1, vbBinaryCompare)
	
	' Insert Day of Month Ordinal (eg st, th, or rd)
	strFormat = Replace(strFormat, "%O", GetDayOrdinal(Day(strDate)), 1, -1, vbBinaryCompare)
	
	' Insert Day of Year
	strFormat = Replace(strFormat, "%j", DatePart("y",strDate), 1, -1, vbBinaryCompare)
	
	' Insert Long Year (4 digit)
	strFormat = Replace(strFormat, "%Y", DatePart("yyyy",strDate), 1, -1, vbBinaryCompare)
	
	' Insert Short Year (2 digit)
	strFormat = Replace(strFormat, "%y", Right(DatePart("yyyy",strDate),2), 1, -1, vbBinaryCompare)
	
	' Insert Weekday as Integer (eg 0 = Sunday)
	strFormat = Replace(strFormat, "%w", DatePart("w",strDate,1), 1, -1, vbBinaryCompare)
	
	' Insert Abbreviated Weekday Name (eg Sun)
	strFormat = Replace(strFormat, "%a", WeekDayName(DatePart("w",strDate,1),True), 1, -1, vbBinaryCompare)
	
	' Insert non-Abbreviated Weekday Name
	strFormat = Replace(strFormat, "%A", WeekDayName(DatePart("w",strDate,1),False), 1, -1, vbBinaryCompare)
	
	' Insert Hour in 24hr format
	str24HourPart = DatePart("h",strDate)
	If Len(str24HourPart) < 2 then str24HourPart = "0" & str24HourPart
	strFormat = Replace(strFormat, "%H", str24HourPart, 1, -1, vbBinaryCompare)
	
	' Insert Hour in 12hr format
	int12HourPart = DatePart("h",strDate) Mod 12
	If int12HourPart = 0 then int12HourPart = 12
	strFormat = Replace(strFormat, "%h", int12HourPart, 1, -1, vbBinaryCompare)
	
	' Insert Minutes
	strMinutePart = DatePart("n",strDate)
	If Len(strMinutePart) < 2 then strMinutePart = "0" & strMinutePart
	strFormat = Replace(strFormat, "%N", strMinutePart, 1, -1, vbBinaryCompare)
	
	' Insert Optional Minutes
	If CInt(strMinutePart) = 0 then
		strFormat = Replace(strFormat, "%n", "", 1, -1, vbBinaryCompare)
	Else
		If CInt(strMinutePart) < 10 then strMinutePart = "0" & strMinutePart
		strMinutePart = ":" & strMinutePart
		strFormat = Replace(strFormat, "%n", strMinutePart, 1, -1, vbBinaryCompare)
	End if
	
	' Insert Seconds
	strSecondPart = DatePart("s",strDate)
	If Len(strSecondPart) < 2 then strSecondPart = "0" & strSecondPart
	strFormat = Replace(strFormat, "%S", strSecondPart, 1, -1, vbBinaryCompare)
	
	' Insert AM/PM indicator
	If DatePart("h",strDate) >= 12 then
		strAMPM = "PM"
	Else
		strAMPM = "AM"
	End If
	
	strFormat = Replace(strFormat, "%P", strAMPM, 1, -1, vbBinaryCompare)
	
	FormatDate = strFormat
	
	'If there is an error output its value
	If err.Number <> 0 then
		Response.Clear()
		Response.Write "ERROR " & err.Number & ": fmcFmtDate - " & err.Description
		Response.Flush()
		Response.End()
	End if

End Function

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


'********************************************************************************
' START STRING OPTIMIZER CLASS
'********************************************************************************

' PURPOSE: 	This is a string optimization class when working with large
'						string concatenations. This class will dramatically speed
'						the process of concatenating strings.
' 			
' ABOUT: 	VB and VBScript have intrinsic array functions such as JOIN that are, 
'					MUCH FASTER at concatenating variant array elements. This class keeps our 
'					string information in an array (its always a Variant anyway, so what's the 
'					difference?) and then when we're finished with all of our concatenations we 
'					have it just do a JOIN on the array with NO DELIMITER so it all comes back 
'					as ONE BIG LONG STRING, in a SINGLE OPERATION.
'
' USAGE:	Set objStrOpt = New StringOptimizer
'					objStrOpt.Reset
'					objStrOpt.Append "String1"
'					objStrOpt.Append "String2"
'					objStrOpt.Append "Stringx"
'					Response.write(objStrOpt.Concat)
Class StringOptimizer
	Dim stringArray,growthRate,numItems

	Private Sub Class_Initialize()
		growthRate = 50: numItems = 0
		ReDim stringArray(growthRate)
	End Sub

	Public Sub Append(ByVal strValue)
		' next line prevents type mismatch error if strValue is null. Performance hit is negligible.
		strValue=strValue & ""
	
		If numItems > UBound(stringArray) Then 
			ReDim Preserve stringArray(UBound(stringArray) + growthRate)
			stringArray(numItems) = strValue: numItems = numItems + 1
		End If

	End Sub

	Public Sub Reset
		Erase stringArray
		Class_Initialize
	End Sub

	Public Function Concat() 
		Redim Preserve stringArray(numItems) 
		Concat = Join(stringArray, "")
	End Function

End Class

'********************************************************************************
' START COMMON FUNCTIONS
'********************************************************************************

Function IsSelected(val,curval)
	If LCase(val) = LCase(curval) Then
		IsSelected = "selected"
	Else
		IsSelected = ""
	End If
End Function

Function FormatMessage(msg)

	Dim strMessage
	strMessage = "<BR>"
	strMessage = strMessage & "<TABLE cellSpacing=0 cellPadding=0 width=100% border=0 ID=Table9999>"
	strMessage = strMessage & "	<TBODY>"
	strMessage = strMessage & "	<TR>"
	strMessage = strMessage & "		<TD colSpan=8 height=8><DIV class=ms-alerttext>" & msg & "</DIV></TD>"
	strMessage = strMessage & "	</TR>"
	strMessage = strMessage & "	</TBODY>"
	strMessage = strMessage & "</TABLE>"
	FormatMessage = strMessage

End Function

' Fixes a string so it can be properly passed to the Db
Function makeString(str)
	Dim newString
	newString = replace(trim(str),"'","''")
	makeString = chr(39) & newString & chr(39)
End Function

' Used for getting the checkbox sql value
Function GetSqlCheckboxValue(val)
	If LCase(val) = "on" Then
		GetSqlCheckboxValue = "Yes"
	Else
		GetSqlCheckboxValue = "No"
	End If
End Function

Function QuoteCleanup(str)
	If StringNotEmptyOrNull(str) Then
		QuoteCleanup = replace(Trim(str),"''","'")
	Else
		QuoteCleanup = ""
	End If
End Function

' Used for wrapping double quotes around a string
Function DblQt(str)
	DblQt = chr(34) & str & chr(34)
End Function

Function StripChars(str,length)
	StripChars = Left(str,length) & "..."
End Function

' Determines if a string is empty or null
Function StringNotEmptyOrNull(strVal)
	If Trim(strVal) & "" <> "" Then
		StringNotEmptyOrNull = True
	Else
		StringNotEmptyOrNull = False
	End If
End Function

Function StringEmptyOrNull(strVal)
	If Trim(strVal) & "" = "" Then
		StringEmptyOrNull = True
	Else
		StringEmptyOrNull = False
	End If
End Function

' Add values to item and subitems of a cookie
Function AddCookie(key,skey,val)
	Dim lclKey
	Dim lclsKey
	
	On Error Resume Next
	
	lclKey = UCase(Trim(key))
	lclsKey = UCase(Trim(skey))	
	if StringNotEmptyOrNull(lclKey) then
		if StringNotEmptyOrNull(lclsKey) then
			Response.Cookies(strCookiePrefix & lclKey)(lclsKey) = val
		else
			response.Cookies(strCookiePrefix & lclKey) = val
		end if
		AddCookie = True
	else
		AddCookie = False
	end if
End Function

Function ExpireCookie(key,when)
	Response.Cookies(strCookiePrefix & Trim(key)).Expires = when
End Function

' Get cookie value based on item and subitem
Function GetCookie(key,skey)
	Dim lclKey
	Dim lclsKey

	On Error Resume Next
	lclKey = UCase(Trim(key))
	lclsKey = UCase(Trim(skey))
	if StringNotEmptyOrNull(lclKey) then
		if StringNotEmptyOrNull(lclsKey) then
			GetCookie = request.Cookies(strCookiePrefix & lclKey)(lclsKey)
		else
			GetCookie = request.Cookies(strCookiePrefix & lclKey)
		end if
	else
		GetCookie = ""
	end if
End Function

' Get the querystring value for given key
Function GetQryString(key)
	Dim lclKey

	lclKey = Trim(key)
	if StringNotEmptyOrNull(lclKey) then
		GetQryString = request.QueryString(lclKey)
	else
		GetQryString = ""
	end if
End Function

' Get the form post value for given key
Function GetFormPost(key)
	if StringNotEmptyOrNull(key) then
		GetFormPost = request.Form(Trim(key))
	else
		GetFormPost = ""
	end if
End Function

' Write a custom message to the IIS web log
Function WriteToLog(strMsg)
	if StringNotEmptyOrNull(strMsg) then
		response.AppendToLog(Trim(strMsg))
		WriteToLog = True	
	else
		WriteToLog = False
	end if
End Function

' Create a custom header 
Function AddToHeader(key,val)
	On Error Resume Next
	if StringNotEmptyOrNull(key) and StringNotEmptyOrNull(val) then
		Response.AddHeader Trim(key),Trim(val)
		AddToHeader = True
	else
		AddToHeader = False
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

Function RemoveSessionVariable(key,blnRemoveAll)
	On Error Resume Next
	If StringNotEmptyOrNull(key) Then
		If IsObject(strSessionPrefix & Session(Trim(key))) Then
			Session.Contents.Remove(strSessionPrefix & Trim(key))
			RemoveSessionVariable = True
		Else
			RemoveSessionVariable = False
		End If
	Else
		If blnRemoveAll Then
			Session.Contents.RemoveAll()
		End If
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

' Create an application (global/cached) variable based on key/value pair
Function AddAppVariable(key,ByVal val)
	On Error Resume Next
	if StringNotEmptyOrNull(key) and StringNotEmptyOrNull(val) then
		Application.Lock()
		If Left(Trim(UCase(key)),4) = "CMS_" Then
			Application(key) = Trim(val)
		Else
			Application(GetSessionVariable("SITE_LOGIN") & "_" & Trim(UCASE(key))) = Trim(val)
		End If		
		Application.UnLock()
	End If
End Function

' Retrieve an application (global/cached) variable based on its key
Function GetAppVariable(key)
	if StringNotEmptyOrNull(key) then
		If Left(Trim(UCase(key)),4) = "CMS_" Then
			GetAppVariable = Application(Trim(UCase(key)))
		Else
			GetAppVariable = Application(GetSessionVariable("SITE_LOGIN") & "_" & Trim(UCase(key)))
		End If
		If GetAppVariable & "" <> "" Then
			GetAppVariable = replace(GetAppVariable,"''","'")
		Else
			GetAppVariable = ""
		End If
	else
		GetAppVariable = ""
	end if
End Function

' Checks for a generic validity of an email address. ( NEED TO GIVE IT CAPABILITIES TO LOOP THROUGH DELIMITED EMAILS ) 
Function IsValidEmail(strEmail)

	Dim blnIsValid
	blnIsValid = True
	
	If Len(strEmail) < 5 Then
		blnIsValid = False
	Else
		If InStr(1, strEmail, "@", 1) < 2 Then
			blnIsValid = False
		Else
			If InStrRev(strEmail, ".") < InStr(1, strEmail, "@", 1) + 2 Then
				blnIsValid = False
			End If
		End If
	End If

	IsValidEmail = blnIsValid
End Function

' PURPOSE:	Sends an email using "CDONTS" with the following parameters. Need to make sure that SMTP is configured in IIS.
'						Also, I have the object defined in the global.asa file. 
' NOTE: 		At this time only 1 email can be passed into each of the email address fields.
' INPUTS:		strFrom - A valid email address ( can only be one email address )
'						strTo - Who this email will be sent to ( can be multiple addresses delimited by comma )
'						strCc - Who this email will be carbon copied to ( can be multiple addresses delimted by comma )
'						strBcc - Who this email will be blind carbon copied to ( can be multiple addresses delimted by comma )
'						strSubject - A description of what the email is about.
'						strBody - The email message to be sent.
'						strType - text (default) or html; specify the format in which to send the email message.
'						strImportance - low, normal, high; level of importance for an email.
Function SendMail(strFrom, strTo, strCc, strBcc, strSubject, strBody, strType, strImportance)

On Error Resume Next
Dim blnSuccess, objMail

' Set default to success
blnSuccess = True

If Not IsValidEmail(trim(strFrom)) Or _
	Not StringNotEmptyOrNull(trim(strTo)) Or _
	Not StringNotEmptyOrNull(trim(strSubject)) Or _
	Not StringNotEmptyOrNull(trim(strBody)) Then
	blnSuccess = False
Else
	Set objMail = Server.CreateObject("CDONTS.NewMail")

	With objMail
		.From = trim(strFrom)
		.To = trim(strTo)
		if StringNotEmptyOrNull(trim(strCc)) then
			if Not IsValidEmail(trim(strCc)) then
				blnSuccess = False
			else
				.Cc = trim(strCc)
			end if
		end if
		if StringNotEmptyOrNull(trim(strBcc)) then
			if Not IsValidEmail(trim(strBcc)) then
				blnSuccess = False
			else
				.Bcc = trim(strBcc)
			end if
		end if
		if StringNotEmptyOrNull(trim(strSubject)) then
			.Subject = trim(strSubject)
		end if
		if StringNotEmptyOrNull(trim(strBody)) then
			.Body = trim(strBody)
		end if
		if StringNotEmptyOrNull(trim(strType)) then
			if lcase(trim(strType)) = "html" then
				.MailFormat = 0
				.BodyFormat = 0
			end if
		end if
		if StringNotEmptyOrNull(trim(strImportance)) then
			select case lcase(trim(strImportance))
				case "low"
					.Importance = 0
				case "normal"
					.Importance = 1
				case "hi"
					.Importance = 2
				case else 'default is normal
					.Importance = 1
			end select
		else
			.Importance = 1
		end if	
		
		.Send
		Set objMail = Nothing
	End With
end if

' check for errors
if err.number <> 0 then
	SendMail = False
else
	SendMail = True
end if

End Function


Function HasErrors(objRs, strType)

	On Error Resume Next

	HasErrors = False
	If Err.Number <> 0 Then
		Session(APPVARNAME & "PathInfo") = request.ServerVariables("PATH_INFO")
		Session(APPVARNAME & "ErrNumber") = Err.Number
		Session(APPVARNAME & "ErrDesc") = Err.Description
		Session(APPVARNAME & "ActionType") = strType
	
		Call CloseRs(objRs)
		Call CloseDb()
		
		HasErrors = True
		Response.Redirect ERRORPAGEPATH
	End If
	
End Function

' PURPOSE: 	Function to call if we want to check for errors in our ASP pages
Function ErrorCheck()
	
	On Error Resume Next
	
	If Err.Number <> 0 then
		response.Write "<form method=" & DblQt("post") & " name=" & DblQt("frmError") & " action=" & DblQt("/SendErrorReport.asp") & ">" & vbCrLf
		response.write "<input type=" & DblQt("hidden") & " name=" & DblQt("ErrorNum") & " value=" & DblQt(Err.Number) & ">" & vbCrLf
		response.write "<input type=" & DblQt("hidden") & " name=" & DblQt("ErrorDesc") & " value=" & DblQt(Err.Description) & ">" & vbCrLf
		response.write "<div class=" & DblQt("error") & ">" & vbCrLf
		response.write "<b>An error occurred while processing the current page.</b><p/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
		response.write "Error Number: " & Err.Number & "<br/>" & vbCrLf
		response.write "Error Description: " & Err.Description & "<p/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & vbCrLf
		response.write "<a href=" & DblQt("javascript:document.frmError.submit()") & ">Send Error Report</a><br/>"
		response.write "</div>"
		response.write "</form>"
		Response.End
	End If		

End Function

Function RaiseError(intNum,strDesc,strSrc,blnLogHistory)
	
	
	if blnLogHistory = True then
		Session("RaiseErrorLog") = Session("RaiseErrorLog") & "::" & intNum & "~" & strDesc & "~" & strSrc
	end if
	
'	Err.Raise intNum
'	Err.Description = strDesc
'	Err.Source = strSrc
'	Response.Write "An Error of the type " & Err.Description & " has been caused by " & Err.Source & "."
	Response.Write "An Error of the type " & strDesc & " has been caused by " & strSrc & "."

End Function

Function GetRaiseErrorLogHistory()

	If Not IsObject(Session("RaiseErrorLog")) then 'is nothing then
'	if Session("RaiseErrorLog") <> "" then
'		Dim arrLog(), intLogLen, intCurrErr, intCurrProp, arrSubItem(), HTML
		HTML = ""
		arrLog = Split(Session("RaiseErrorLog"),"::")
		For intCurrErr = LBound(arrLog) To UBound(arrLog)
			arrSubItem = Split(arrLog(intCurrErr),"~")
			HTML = HTML & "<ol>"
			For intCurrProp = LBound(arrSubItem) To UBound(arrSubItem)
				HTML = HTML & "<li>" & arrSubItem(intCurrProp)
			Next

			HTML = HTML & "</ol>"
		Next
		response.write HTML
		GetRaiseErrorLogHistory = HTML
	else
		GetRaiseErrorLogHistory = ""
	end if
		
End Function

Function ClearRaiseErrorLogHistory()
	Session.Contents.Remove("RaiseErrorLog")
'	Set Session("RaiseErrorLog") = Nothing
End Function

'PURPOSE: Expires a page. Call this function from the top of an un-cached page.
Sub ExpirePage(intDuration)
	Response.Expires = intDuration
End Sub

' CREATED BY: 	BRIAN GAINES
' CREATED ON: 	02/27/2003
' PURPOSE: 			This function servers the purpose of displaying HTML
'								comments while applying different styles.
' INPUTS:				comment - The text you want to make a comment
'								commentType - This is the type (or style) of comment to display
' COPYRIGHT: 		Feel free to use this function within your own applications.
Sub HTMLComment(comment,commentType)

	If comment <> "" Then		
		Response.Write(vbCrLf & "<!-- ")
		If commentType = 1 Then 		' HEADER START BLOCK
			Response.Write("START: " & comment)
		ElseIf commentType = 2 Then		' HEADER END BLOCK
			Response.Write("END: " & comment)
		ElseIf commentType = 3 Then 	' REGULAR COMMENT BLOCK
			Response.Write("COMMENT: " & comment)
		Else
			Response.Write(comment)		' GENERIC FORMAT
		End If
		Response.Write(" -->" & vbCrLf)
	End If
	
End Sub

Function KillUserLogin(strRedirect)
	Session.Abandon()
	Response.Redirect strRedirect
End Function

Function AuthenticateUser(strUserId,strUserPwd,strCookie)

		If strUserId="" or strUserPwd="" Then
			AuthenticateUser = "MISSING CREDENTIALS"
			Response.redirect("login.asp?msg=You+must+enter+a+User+Id+and+Password!")
		End If

		Dim sqlLoginCheck
		Dim rsVerifyLogin
		sqlLoginCheck = "SELECT * FROM tblUsers WHERE email=" & makeString(strUserId)
		Call OpenRS(rsVerifyLogin)
		'rsVerifyLogin.Open sqlLoginCheck, GetConnection
		Call ExecuteSql(sqlLoginCheck,rsVerifyLogin)
	
		if Err.Number <> 0 then
			response.write "An error occurred while trying to authenticate the user"
			response.End()
		end if	
	
		With rsVerifyLogin
			If .BOF and .EOF Then
				If session("userCount") <> "" Then
					session("userCount") = session("userCount") + 1
				Else
					session("userCount") = 1
				End If
				AuthenticateUser = "INVALID LOGIN"
				Exit Function
			Else
				.movefirst
				If rsVerifyLogin("lockout") = True Then
					AuthenticateUser = "USER LOCKED OUT"
					Exit Function
				Else
					If session("__Email") <> userID Then
						session("__Email") = userId
						session("memberCount") = 1
					Else
						session("memberCount") = session("memberCount") + 1
					End If
		
					' decrypt the password here
					If DecryptText(rsVerifyLogin("pw")) <> userPassword Then
						AuthenticateUser = "INCORRECT PASSWORD"
						Exit Function
					Else
	
						' Remove login session variables
						Session.Contents.Remove("userCount")
						Session.Contents.Remove("memberCount")	
	
						session("__UserId") = rsVerifyLogin("id")
						session("__Email") = rsVerifyLogin("email")
						session(APPVARNAME & "RoleId") = rsVerifyLogin("Role")
						session("__UserFullName") = rsVerifyLogin("FullName")
						session("__position") = rsVerifyLogin("empposition")
						session("__company") = rsVerifyLogin("Company")
						session("__LoginAsCompany") = session("company")
						session("__address") = rsVerifyLogin("Address")
						session("__city") = rsVerifyLogin("City")
						session("__state") = rsVerifyLogin("State")
						session("__zip") = rsVerifyLogin("Zip")
						session("__phone") = rsVerifyLogin("Phone")
						session("__mobile") = rsVerifyLogin("Mobile")
						session("__fax") = rsVerifyLogin("Fax")
						session("__lastlogin") = rsVerifyLogin("LastLogin")
						session("__imagelogo") = rsVerifyLogin("imagename")
						
						session("__UserAuthenticated") = True
						
						If Session("PersistLogin") = True Then
							Response.Cookies(strCookie)("UserName") = session("__Email")
							Response.Cookies(strCookie)("Password") = EncryptText(userPassword)
							Response.Cookies(strCookie).Expires = now() + 7
							Session.Contents.Remove("__PersistLogin")
						Else
							Response.Cookies(strCookie).Expires = now() - 1
						End If
		
						AuthenticateUser = "AUTHENTICATED"
					End If
				End If
			End If
		End With
	
		Call CloseRS(rsVerifyLogin)
End Function

Sub CheckLoginAttempts(strAppName)

  If session("memberCount") => 3 Then
		Response.write "<img src='" & Application(strAppName & "AppImagePath") & "/stop.gif' border='0' alt='Login Error'>&nbsp;"
		Response.write("You have incorrectly tried to access the " & Application("__AppCompanyName") & " member area 3 or more times.<BR>")
		Response.write("The account " & session("__Email") & " has been locked.<BR>")
		Response.write("Contact the administrator to unlock the account.<BR>")

		Call OpenDB()
		Dim updSql
		updSql = "UPDATE tblUsers SET lockout = Yes WHERE email = '" & session("__Email") & "'"
		dbConn.Execute updSql
		Call CloseDB()
		Response.End()
  Elseif session("userCount") => 5 Then
		Response.write "<img src='" & Application(APPVARNAME & "AppImagePath") & "/stop.gif' border='0' alt='Login Error'>&nbsp;"
		Response.write("You have incorrectly tried to access the " & Application(APPVARNAME & "AppCompanyName") & " member area 5 or more times.<BR>")        
		Response.End()
  End If

End Sub

Function EncryptText(strText)

	Dim CharHexSet, intStringLen, strTemp, strRAW, i, intKey, intOffSet
	Randomize Timer

	intKey = Round((RND * 1000000) + 1000000)   		'##### Key Bitsize
	intOffSet = Round((RND * 1000000) + 1000000)   	'##### KeyOffSet Bitsize
	
		If IsNull(strText) = False Then
			strRAW = strText
			intStringLen = Len(strRAW)
					
			For i = 0 to intStringLen - 1
				strTemp = Left(strRAW, 1)
				strRAW = Right(strRAW, Len(strRAW) - 1)
				CharHexSet = CharHexSet & Hex(Asc(strTemp) * intKey) & Hex(intKey)
			Next
			
			EncryptText = CharHexSet & "|" & Hex(intOffSet + intKey) & "|" & Hex(intOffSet)
		Else
			EncryptText = ""
		End If
		
End Function

Function DecryptText(strText)

	Dim strRAW, arHexCharSet, i, intKey, intOffSet, strRawKey, strHexCrypData
	
	strRawKey = Right(strText, Len(strText) - InStr(strText, "|"))
	intOffSet = Right(strRawKey, Len(strRawKey) - InStr(strRawKey,"|"))
	intKey = HexConv(Left(strRawKey, InStr(strRawKey, "|") - 1)) - HexConv(intOffSet)
	strHexCrypData = Left(strText, Len(strText) - (Len(strRawKey) + 1))

	arHexCharSet = Split(strHexCrypData, Hex(intKey))
		
	For i=0 to UBound(arHexCharSet)
		strRAW = strRAW & Chr(HexConv(arHexCharSet(i))/intKey)
	Next
		
	DecryptText = strRAW
	
End Function

Sub NoCache
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
End Sub

Private Function HexConv(hexVar)
Dim hxx, hxx_var, multiply          
        IF hexVar <> "" THEN
             hexVar = UCASE(hexVar)
             hexVar = StrReverse(hexVar)
             DIM hx()
             REDIM hx(LEN(hexVar))
             hxx = 0
             hxx_var = 0
             FOR hxx = 1 TO LEN(hexVar)
                  IF multiply = "" THEN multiply = 1
                  hx(hxx) = mid(hexVar,hxx,1)
                  hxx_var = (get_hxno(hx(hxx)) * multiply) + hxx_var
                  multiply = (multiply * 16)
             NEXT
             hexVar = hxx_var
             HexConv = hexVar
        END IF
End Function
   
Private Function get_hxno(ghx)
        If ghx = "A" Then
             ghx = 10
        ElseIf ghx = "B" Then
             ghx = 11
        ElseIf ghx = "C" Then
             ghx = 12
        ElseIf ghx = "D" Then
             ghx = 13
        ElseIf ghx = "E" Then
             ghx = 14
        ElseIf ghx = "F" Then
             ghx = 15
        End If
        get_hxno = ghx
End Function

Function CharCount(strSource, strChar)
    Dim intPos, intCount
    intCount = 0
    intPos = InStr(strSource, strChar)
    While intPos
        intCount = intCount + 1
        intPos = InStr(intPos + 1, strSource, strChar)
    Wend
    CharCount = intCount
End Function

Function PageFormat(strDbValue)
	Select Case strDbValue
		Case "empty", "12/31/1899", "12/31/1899 12:00:00 AM", "12/31/1899 12:00 AM"
			PageFormat = ""
		Case Else
			PageFormat = strDbValue
	End Select
End Function

'used to format the checkbox by determining the value from the database
Function CkBoxPageFormat(blnDbValue)
	If blnDbValue = True Then
		CkBoxPageFormat = " checked"
	ElseIf blnDbValue = False Then
		CkBoxPageFormat = ""
	End IF
End Function

'used to format the checkbox to the original setting after the form was posted
'i.e If record already exists and want to display original values
Function CkBoxPagePostBack(strValue)
	If strValue = "on" Then
		CkBoxPagePostBack = " checked"
	Else
		CkBoxPagePostBack = ""
	End If
End Function

'used to format the value from the checkbox when it is saved to the database
Function GetSqlCheckboxValue(strValue)
	If strValue = "on" Then
		GetSqlCheckboxValue = 1
	Else
		GetSqlCheckboxValue = 0
	End If
End Function

Function IsChecked(strValue)
	If strValue Then
		IsChecked = "checked"
	Else
		IsChecked = ""
	End If
End Function

Function GetCheckboxValue(strValue,IsReport)

	Select Case strValue
		Case "on","off"
			If strValue = "on" Then
				GetCheckboxValue = "1"
			Else
				GetCheckboxValue = "0"
			End If
		Case Else
			If strValue = -1 Then
				If IsReport Then
					GetCheckboxValue = "Active"
				Else
					GetCheckboxValue = "checked"
				End If
			Else
				If IsReport Then
					GetCheckboxValue = "Inactive"
				Else
					GetCheckboxValue = ""
				End If
			End If
	End Select
End Function

Function FormatActiveFlag(flag)
	If (flag) Then 
		FormatActiveFlag = "Active"
	Else
		FormatActiveFlag = "<FONT STYLE=""COLOR:Red"">Inactive</FONT>"
	End If
End Function

Sub PageRedirect(path)
	Response.Redirect(path)
End Sub

Function GetPageMode(strMode)
	If strMode <> "" Then
		GetPageMode = UCase(strMode)
	Else
		GetPageMode = "NEW"
	End If
End Function


Function GetExpirationDate(strDate)

	If StringNotEmptyOrNull(strDate) Then
		If strDate <> "[Never Expire]" Then
			GetExpirationDate = "#" & strDate & "#"
		Else
			GetExpirationDate = "#12/31/2030 11:59:59 PM#"
		End If
	End If

End Function

Function ProperCase(strText)

	'Results will appear as SomeTextLikeThis.txt

	Dim arrName
	Dim intCounter
	Dim tmpName
	Dim tmpStr

	'--------------------------
	' Split filename on space
	'--------------------------
	arrName = Split(strText," ")

	'----------------------------------
	' If there is a space in file name
	'----------------------------------
	
	If UBound(arrName) > 0 Then

		'-------------------------
		' Loop through each item
		'-------------------------

		For intCounter = 0 To UBound(arrName)

			'------------------------------------------
			' Trim off any leading/trailing spaces
			'------------------------------------------
			tmpName = Trim(arrName(intCounter))
			
			'-----------------------------------------------
			' Capitalize 1st letter and make rest lowercase
			'-----------------------------------------------			
			tmpStr = tmpStr & Trim(UCase(Left(tmpName,1))) & Trim(Right(tmpName,Len(tmpName)-1))
		Next
		ProperCase = tmpStr
	Else

		'-----------------------------------------------
		' Capitalize 1st letter and make rest lowercase
		'-----------------------------------------------			

		ProperCase = Trim(UCase(Left(Trim(strText),1)) & Trim(LCase(Right(Trim(strText),Len(Trim(strText))-1))))
	End If

End Function


Function ShrinkText(strText)

	'Results will change Some Text Like This.txt to sometextlikethis.txt

	Dim arrName
	Dim intCounter
	Dim tmpName
	Dim tmpStr

	'--------------------------
	' Split filename on space
	'--------------------------
	arrName = Split(strText," ")

	'----------------------------------
	' If there is a space in file name
	'----------------------------------
	
	If UBound(arrName) > 0 Then

		'-------------------------
		' Loop through each item
		'-------------------------

		For intCounter = 0 To UBound(arrName)

			'------------------------------------------
			' Trim off any leading/trailing spaces
			'------------------------------------------
			tmpName = Trim(arrName(intCounter))
			
			'-----------------------------------------------
			' Capitalize 1st letter and make rest lowercase
			'-----------------------------------------------			
			tmpStr = tmpStr & Trim(LCase(tmpName))
		Next
		ShrinkText = tmpStr
	Else

		'-----------------------------------------------
		' Make current text lowercase
		'-----------------------------------------------			

		ShrinkText = Trim(LCase(strText))
	End If

End Function


' CREATED BY:	BRIAN GAINES
' PURPOSE: 		Builds the combo box with list of values from database and selects the appropriate option
'							based on the valToMatch value passed into this function.			
' INPUTS: 		rsSelect - Recordset object that will populate the combo box
'							optName - Name to give to combo box
'							valToMatch - The value to match up with the recordset's optValue
Function BuildComboBox(rsSelect,optName,valToMatch)

	Dim strList	

	'NOTE: 	The recordset option value must represented by the first item => rsSelect(0).Value,
	' 		whereas the option text must be represented by the second item => rsSelect(1).Value

	' build up the combo box list. 
	strList = "<SELECT NAME=" & DblQt(optName) & ">" & vbCRLF
	With rsSelect
		.MoveFirst
		do while not .EOF
			strList = strList & "<OPTION VALUE=" & DblQt(rsSelect(0)) & setComboOption(rsSelect(0),valToMatch,1) & ">" & rsSelect(1) & "</OPTION>" & vbCRLF
			.MoveNext
		loop
	End With
	strList = strList & "</SELECT>" & vbCRLF

	' return the combo box list to the caller
	BuildComboBox = strList

End Function

' PURPOSE:	Sets the appropriate value depending on the type of control is calling this control.
' INPUTS:		optValue - This is the value from the lookup table
'						DeltaValue - This is the value that we should select the option value on.
'						OptionRadio - 0 or 1 to determine if it is a ComboBox or Checkbox control				
Function setComboOption (optValue,DeltaValue,OptionRadio)
	' if there is a match then select/check it	
	if optValue = DeltaValue Then
		Select Case OptionRadio
			Case 1	' select
				setComboOption = " selected "
			Case 2	' checkbox
				setComboOption = " checked "
		End Select
	End if
End Function


%>
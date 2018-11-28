<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/objects/cLogin.asp" -->
<%
Dim PAGE_IMAGE, EDIT_PAGE, REPORT_PAGE, PAGE_TITLE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_content.gif"
PAGE_TITLE = "User Management"
EDIT_PAGE = "user.asp"
REPORT_PAGE = "user-report.asp"
IS_REPORT_PAGE = False

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function AddUser(Login, Pwd, Email, FullName, Phone, Address1, Address2, City, StateCode, Zip, Comments, Expire, IsAdmin, UserId)
	Dim oLogin
	Set oLogin = New cLogin
	With oLogin
		.Login = Login
		.Email = Email
		.Password = Pwd
		.FullName = FullName
		.Phone = Phone
		.Address1 = Address1
		.Address2 = Address2
		.City = City
		.StateCode = StateCode
		.Zip = Zip
		.Comments = Comments
		.Photographer = PHOTOGRAPHER_FNAME
		.Expire = Expire
		.IsAdmin = IsAdmin
		.AddUser()
		intUserId = .ID
		AddUser = Not .IsError
	End With
	Set oLogin=Nothing
End Function

Function UpdateUser(Login, Pwd, Email, FullName, Phone, Address1, Address2, City, StateCode, Zip, Comments, Expire, IsAdmin, UserId)
	Dim oLogin
	Set oLogin = New cLogin
	With oLogin
		.ID = userId
		.Login = Login
		.Email = Email
		.Password = Pwd
		.FullName = FullName
		.Phone = Phone
		.Address1 = Address1
		.Address2 = Address2
		.City = City
		.StateCode = StateCode
		.Zip = Zip
		.Comments = Comments
		.Photographer = PHOTOGRAPHER_FNAME
		.Expire = Expire
		.IsAdmin = IsAdmin
		.UpdateUser()
		UpdateUser = Not .IsError
	End With
	Set oLogin = Nothing
End Function

'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim intUserId
Dim strLogin
Dim strPwd
Dim strEmail
Dim strFullName
Dim strPhone
Dim strAddress1
Dim strAddress2
Dim strCity
Dim strStateCode
Dim strZip
Dim strComments
Dim strExpire
Dim intIsAdmin

'----------------------------------------------------------------------------------------
' PAGE RENDER LOGIC
'----------------------------------------------------------------------------------------
If GetFormPost("hidMode") = "save" Then
	strMode = "edit"	
	intUserId = GetFormPost("hidUserId")
	strLogin = GetFormPost("txtLogin")
	strEmail = GetFormPost("txtEmail")
	strPwd = GetFormPost("txtPassword")
	strFullName = GetFormPost("txtFullName")
	strPhone = GetFormPost("txtPhone")
	strAddress1 = GetFormPost("txtAddress1")
	strAddress2 = GetFormPost("txtAddress2")	
	strCity = GetFormPost("txtCity")
	strStateCode = GetFormPost("txtState")
	strZip = GetFormPost("txtZip")		
	strComments = QuoteCleanup(GetFormPost("taComments"))
	strExpire = GetFormPost("txtExpire")
	intIsAdmin = GetSqlCheckboxValue(GetFormPost("ckIsAdmin"))

	if StringNotEmptyOrNull(intUserId) Then
		blnSuccess = UpdateUser(strLogin, strPwd, strEmail, strFullName, strPhone, strAddress1, strAddress2, strCity, strStateCode, strZip, strComments, strExpire, intIsAdmin, intUserId)
	Else
		blnSuccess = AddUser(strLogin, strPwd, strEmail, strFullName, strPhone, strAddress1, strAddress2, strCity, strStateCode, strZip, strComments, strExpire, intIsAdmin, intUserId)
	End If	

	If blnSuccess Then
		PageRedirect(REPORT_PAGE)
	Else
		displayMessage = "An error occurred while trying to save information to the database."
	End If
Else
	intUserId = GetQryString("id")

	If StringNotEmptyOrNull(intUserId) Then
		Dim collLogin, oLogin
		Set oLogin = New cLogin
		Set collLogin = New cLogin
		collLogin.ID = intUserId
		Call collLogin.GetUserById()
		For Each oLogin In collLogin.Logins.Items
			intUserId = oLogin.ID
			strLogin = oLogin.Login
			strPwd = oLogin.Password
			strEmail = oLogin.Email
			strFullName = oLogin.FullName
			strPhone = oLogin.Phone
			strAddress1 = oLogin.Address1
			strAddress2 = oLogin.Address2
			strCity = oLogin.City
			strStateCode = oLogin.StateCode
			strZip = oLogin.Zip
			strComments = oLogin.Comments
			strPhotographer = oLogin.Photographer
			strExpire = oLogin.Expire
			intIsAdmin = oLogin.IsAdmin
			Set oLogin = Nothing
		Next
		Set collLogin = Nothing
	
	Else
		strExpire = FormatDate(Now()+DAYS_FROM_NOW_TO_EXPIRE_ACCT,"%m/%d/%Y")
	End If
End If
%>
<HTML>
<HEAD>
<TITLE><%=PAGE_TITLE%></TITLE>
<script src="/cms/lib/__dom.js" type="text/javascript"></script>
<script src="/cms/lib/__calendar2.js" type="text/javascript"></script>
<script>
function checkForm() {
	var frm = document.forms['form1'];
	var mode = frm.hidMode;
	mode.value = 'save';
	frm.submit();
}
</script>
<!--#include virtual="/cms/styles.asp"-->
</HEAD>
<BODY>
<form name="form1" id="form1" method="post" action="<%=EDIT_PAGE%>">
<input type="hidden" name="hidUserId" id="hidUserId" value="<%=intUserId%>">
<input type="hidden" name="hidMode" id="hidMode" value="<%=strMode%>">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" style="border-bottom-style:solid;border-width:1;border-color:#666;">
	<!--#include virtual="/cms/common/__header.asp"-->
	<!--#include virtual="/cms/common/__titlebar.asp" -->
	<!--#include virtual="/cms/common/__begin_bodywrap.asp"-->
	<table class="table-container" width="<%=CMS_PAGE_WIDTH%>">
		<tr>
			<th align="left" width="200">Full Name</th>
			<td align="left" width="550"><input type="text" size="50" id="txtFullName" name="txtFullName" value="<%=strFullName%>"></td>
		</tr>			
<!--		<tr>
			<th align="left" width="200">Email Address<em>*</em></th>
			<td align="left" width="550"><input type="text" size="50" id="txtEmail" name="txtEmail" value="<%=strEmail%>"></td>
		</tr>													
		<tr>
			<th align="left" width="200">Phone Number<em>*</em></th>
			<td align="left" width="550"><input type="text" size="50" id="txtPhone" name="txtPhone" value="<%=strPhone%>"></td>
		</tr>
		<tr>
			<th align="left" width="200">Address 1<em>*</em></th>
			<td align="left" width="550"><input type="text" size="50" id="txtAddress1" name="txtAddress1" value="<%=strAddress1%>"></td>
		</tr>								
		<tr>
			<th align="left" width="200">Address 2</th>
			<td align="left" width="550"><input type="text" size="50" id="txtAddress2" name="txtAddress2" value="<%=strAddress2%>"></td>
		</tr>							
		<tr>
			<th align="left" width="200">City, State Zip<em>*</em></th>
			<td align="left" width="550">
				<input type="text" size="26" id="txtCity" name="txtCity" value="<%=strCity%>">, 
				<input type="text" size="5" id="txtState" name="txtState" value="<%=strStateCode%>"> 
				<input type="text" size="10" id="txtZip" name="txtZip" value="<%=strZip%>">
			</td>
		</tr>								
-->		
		<tr>
			<th align="left" width="200">Login</th>
			<td align="left" width="550"><input type="text" size="50" id="txtLogin" name="txtLogin" value="<%=strLogin%>"></td>
		</tr>								
		<tr>
			<th align="left" width="200">Password</th>
			<td align="left" width="550"><input type="text" size="50" id="txtPassword" name="txtPassword" value="<%=strPwd%>"></td>
		</tr>								
<!--
		<tr>
			<th align="left" width="200">Comments<em>*</em></th>
			<td align="left" width="550"><textarea rows="5" cols="60" id="taComments" name="taComments"><%=strComments%></textarea></td>
		</tr>
-->		
		<tr>
			<th align="left" width="200">Expire Account</th>
			<td align="left" width="550"><input type="text" size="25" id="txtExpire" name="txtExpire" value="<%=strExpire%>">&nbsp;<i>(mm/dd/yyyy)</i>
<!--			<a href="javascript:;" onClick="calExpire.popup();"><img src="<%=CMS_IMAGE_PATH%>/calendar.gif" border="0" alt="Click to Pick a Expiration Date"></a>-->
			</td>
		</tr>
		<%
		Response.write "<SCR" & "IPT>" & vbcrlf
		Response.write "var domExpire = findDOM('txtExpire');" & vbcrlf
		Response.write "domExpire.value='" & strExpire & "';" & vbcrlf
		Response.write "domExpire.disabled=false;" & vbcrlf							
		Response.write "</SCR" & "IPT>" & vbcrlf
		%>		

		<tr>
			<th align="left" width="200">Is Admin</th>
			<td align="left" width="550"><input type="checkbox" id="ckIsAdmin" name="ckIsAdmin" <%=IsChecked(intIsAdmin)%>></td>
		</tr>
	</table>
	<!--#include virtual="/cms/common/__end_bodywrap.asp"-->
</table>
<!--#include virtual="/cms/common/__footer.asp"-->
<script language="JavaScript">
<!-- // create calendar object(s) just after form tag closed
	var calExpire = new calendar2(domExpire);
	calExpire.year_scroll = true;
	calExpire.time_comp = false;
//-->
</script>
</form>

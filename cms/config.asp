<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<%
Dim PAGE_IMAGE, EDIT_PAGE, REPORT_PAGE, PAGE_TITLE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_listset.gif"
PAGE_TITLE = "Configuration Management"
EDIT_PAGE = "config.asp"
REPORT_PAGE = "config-report.asp"
IS_REPORT_PAGE = False

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function AddConfig(configKey,configValue,configDesc,intConfigId)
	Dim oConfig
	Set oConfig = New cConfig
	With oConfig
		.ConfigKey = configKey
		.ConfigValue = configValue
		.ConfigDesc = configDesc
		.AddConfig()
		intConfigId = .ID
		AddConfig = Not .IsError
	End With
	Set oConfig=Nothing
End Function

Function UpdateConfig(id,configKey,configValue,configDesc)
	Dim oConfig
	Set oConfig = New cConfig
	With oConfig
		.ID = id
		.ConfigKey = configKey
		.ConfigValue = configValue
		.ConfigDesc = configDesc
		.UpdateConfig()
		UpdateConfig = Not .IsError
	End With
	Set oConfig = Nothing
End Function

'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim intConfigId
Dim strConfigKey
Dim strConfigValue
Dim strConfigDesc

'----------------------------------------------------------------------------------------
' PAGE RENDER LOGIC
'----------------------------------------------------------------------------------------
If GetFormPost("hidMode") = "save" Then
	strMode = "edit"	
	intConfigId = GetFormPost("hidConfigId")
	strConfigKey = GetFormPost("txtConfigKey")
	strConfigValue = QuoteCleanup(GetFormPost("txtConfigValue"))
	strConfigDesc = QuoteCleanup(GetFormPost("taConfigDesc"))
	
	If StringNotEmptyOrNull(intConfigId) Then
		blnSuccess = UpdateConfig(intConfigId,strConfigKey,strConfigValue,strConfigDesc)		
	Else
		blnSuccess = AddConfig(strConfigKey,strConfigValue,strConfigDesc, intConfigId)
	End If
	
	If blnSuccess Then
		PageRedirect(REPORT_PAGE)
	Else
		displayMessage = "An error occurred while trying to save information to the database."
	End If
Else
	intConfigId = GetQryString("id")
	If StringNotEmptyOrNull(intConfigId) Then
		Set oConfig = New cConfig
		Set collConfig = New cConfig
		collConfig.ID = intConfigId
		Call collConfig.GetConfigInfoById()
		For Each oConfig In collConfig.Configs.Items		
			intConfigId = oConfig.ID
			strConfigKey = oConfig.ConfigKey
			strConfigValue = oConfig.ConfigValue
			strConfigDesc = oConfig.ConfigDesc
			Set oConfig = Nothing
		Next
		Set collConfig = Nothing
	End If
End If
%>
<HTML>
<HEAD>
<TITLE><%=PAGE_TITLE%></TITLE>
<script src="/cms/lib/__dom.js" type="text/javascript"></script>
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
<input type="hidden" id="hidConfigId" name="hidConfigId" value="<%=intConfigId%>">
<input type="hidden" name="hidMode" id="hidMode" value="<%=strMode%>">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" style="border-bottom-style:solid;border-width:1;border-color:#666;">
	<!--#include virtual="/cms/common/__header.asp"-->
	<!--#include virtual="/cms/common/__titlebar.asp" -->
	<!--#include virtual="/cms/common/__begin_bodywrap.asp"-->
	<table class="table-container" width="<%=CMS_PAGE_WIDTH%>">
		<tr>
			<th align="left" width="200">Config Key<em>*</em></th>
			<td align="left" width="550"><input type="text" size="50" id="txtConfigKey" name="txtConfigKey" value="<%=strConfigKey%>"></td>
		</tr>
		<tr>
			<th align="left" width="200">Config Value<em>*</em></th>
			<td align="left" width="550"><input type="text" size="50" id="txtConfigValue" name="txtConfigValue" value="<%=strConfigValue%>"></td>
		</tr>
		<tr>
			<th align="left" width="200">Config Description</th>
			<td align="left" width="550"><textarea cols="80" rows="10" wrap="virtual" id="taConfigDesc" name="taConfigDesc"><%=strConfigDesc%></textarea>
		</tr>
		<%=RenderRequiredFieldsMessageRow%>
	</table>									
	<!--#include virtual="/cms/common/__end_bodywrap.asp"-->
</table>
<!--#include virtual="/cms/common/__footer.asp"-->
</form>
</BODY>
</HTML>

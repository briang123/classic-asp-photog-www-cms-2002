<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/objects/cMetaData.asp" -->
<!-- #include virtual="/objects/cPageInfo.asp" -->
<%
Dim PAGE_IMAGE, EDIT_PAGE, REPORT_PAGE, PAGE_TITLE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_listset.gif"
PAGE_TITLE = "Meta Tags for Search Engine"
EDIT_PAGE = "meta.asp"
REPORT_PAGE = "meta-report.asp"
IS_REPORT_PAGE = False

Public Function GetPageInfo(ctlName,val)

	Dim tempStr
	
	tempStr = "<select name=""" & ctlName & """ id=""" & ctlName & """>"
'	tempStr = tempStr & "<option value=""0"""
'	If StringEmptyOrNull(val) Then
'		tempStr = tempStr & " selected"
'	End If
'	tempStr = tempStr & ">---SELECT ONE---</option>"

	Dim collPageInfo, oPageInfo
	Set oPageInfo = New cPageInfo
	Set collPageInfo = New cPageInfo
	Call collPageInfo.GetPageInfo()
	For Each oPageInfo In collPageInfo.PageInfo.Items
		tempStr = tempStr & "<option value=""" & oPageInfo.ID & """"
		If oPageInfo.ID = val Then
			tempStr = tempStr & " selected>"
		Else
			tempStr = tempStr & ">"
		End If
		tempStr = tempStr & QuoteCleanup(oPageInfo.WebPage) & "</option>"
	Next
	Set oPageInfo = Nothing
	Set collPageInfo = Nothing
	tempStr = tempStr & "</select>"
	
	GetPageInfo = tempStr
	
End Function

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function AddMetaData(webPage,metaKeywords,metaDesc,intMetaId)
	Dim oMeta
	Set oMeta = New cMetaData
	With oMeta
		.WebPageId = webPage
		.MetaKeywords = metaKeywords
		.MetaDescription = metaDesc
		.AddMetaData()
		intMetaId = .ID
		AddMetaData = Not .IsError
	End With
	Set oMeta=Nothing
End Function

Function UpdateMetaData(id,webPage,metaKeywords,metaDesc)
	Dim oMeta
	Set oMeta = New cMetaData
	With oMeta
		.ID = id
		.WebPageId = webPage
		.MetaKeywords = metaKeywords
		.MetaDescription = metaDesc
		.UpdateMetaData()
		UpdateMetaData = Not .IsError
	End With
	Set oMeta = Nothing
End Function

'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim intMetaId
Dim intWebPageId
Dim strMetaKeywords
Dim strMetaDescription

'----------------------------------------------------------------------------------------
' PAGE RENDER LOGIC
'----------------------------------------------------------------------------------------
If GetFormPost("hidMode") = "save" Then
	strMode = "edit"	
	intMetaId = GetFormPost("hidMetaId")
	intWebPageId = GetFormPost("selPageInfo")
	strMetaKeywords = QuoteCleanup(GetFormPost("taMetaKeywords"))
	strMetaDescription = QuoteCleanup(GetFormPost("taMetaDescription"))
	
	If StringNotEmptyOrNull(intMetaId) Then
		blnSuccess = UpdateMetaData(intMetaId,intWebPageId,strMetaKeywords,strMetaDescription)		
	Else
		blnSuccess = AddMetaData(intWebPageId,strMetaKeywords,strMetaDescription,intMetaId)
	End If
	
	If blnSuccess Then
		PageRedirect(REPORT_PAGE)
	Else
		displayMessage = "An error occurred while trying to save information to the database."
	End If
Else
	intMetaId = GetQryString("id")
	If StringNotEmptyOrNull(intMetaId) Then
		Set oMeta = New cMetaData
		Set collMeta = New cMetaData
		collMeta.ID = intMetaId
		Call collMeta.GetMetaDataById()
		For Each oMeta In collMeta.MetaData.Items		
			intMetaId = oMeta.ID
			intWebPageId = oMeta.WebPageId
			strMetaKeywords = oMeta.MetaKeywords
			strMetaDescription = oMeta.MetaDescription
			Set oMeta = Nothing
		Next
		Set collMeta = Nothing
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
<input type="hidden" id="hidMetaId" name="hidMetaId" value="<%=intMetaId%>">
<input type="hidden" name="hidMode" id="hidMode" value="<%=strMode%>">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" style="border-bottom-style:solid;border-width:1;border-color:#666;">
	<!--#include virtual="/cms/common/__header.asp"-->
	<!--#include virtual="/cms/common/__titlebar.asp" -->
	<!--#include virtual="/cms/common/__begin_bodywrap.asp"-->
	<table class="table-container" width="<%=CMS_PAGE_WIDTH%>">
		<tr>
			<th align="left" width="200">Web Page</th>
			<td align="left" width="550"><%=GetPageInfo("selPageInfo",intWebPageId)%></td>
		</tr>
		<tr>
			<th align="left" width="200">Search Keywords</th>
			<td align="left" width="550"><textarea cols="80" rows="5" wrap="virtual" id="taMetaKeywords" name="taMetaKeywords"><%=strMetaKeywords%></textarea>
		</tr>
		<tr>
			<th align="left" width="200">Search Description</th>
			<td align="left" width="550"><textarea cols="80" rows="7" wrap="virtual" id="taMetaDescription" name="taMetaDescription"><%=strMetaDescription%></textarea>
		</tr>
	</table>									
	<!--#include virtual="/cms/common/__end_bodywrap.asp"-->
</table>
<!--#include virtual="/cms/common/__footer.asp"-->
</form>
</BODY>
</HTML>

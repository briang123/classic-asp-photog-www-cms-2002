<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/lib/__pagepermissions.asp" -->
<!-- #include virtual="/objects/cAbout.asp" -->
<%
Dim PAGE_IMAGE, EDIT_PAGE, REPORT_PAGE, PAGE_TITLE, IS_REPORT_PAGE
PAGE_IMAGE = "rect_content.gif"
PAGE_TITLE = "About Me"
EDIT_PAGE = "about.asp"
REPORT_PAGE = "about-report.asp"
IS_REPORT_PAGE = False

'----------------------------------------------------------------------------------------
' REFACTORED FUNCTION CALLS TO CLASS OBJECTS
'----------------------------------------------------------------------------------------
Function AddAboutText(aboutText,lngAboutId)
	Dim oAbout
	Set oAbout = New cAbout
	With oAbout
		.AboutText = aboutText
		.AddAboutText()
		lngAboutId = .ID
		AddAboutText = Not .IsError
	End With
	Set oAbout = Nothing
End Function

Function UpdateAboutText(id,aboutText)
	Dim oAbout
	Set oAbout = New cAbout
	With oAbout
		.ID = id
		.AboutText = aboutText
		.UpdateAboutText()
		UpdateAboutText = Not .IsError
	End With
	Set oAbout = Nothing
End Function

'----------------------------------------------------------------------------------------
' VARIABLE DECLARATIONS
'----------------------------------------------------------------------------------------
Dim lngAboutId
Dim strAboutText

'----------------------------------------------------------------------------------------
' PAGE RENDER LOGIC
'----------------------------------------------------------------------------------------
If GetFormPost("hidMode") = "save" Then
	strMode = "edit"	
	lngAboutId = GetFormPost("hidAboutId")
	
	//die(Request.Form("taAboutText"))
    //die(Request.Form("editorContent"))
	
	strAboutText = QuoteCleanup(GetFormPost("editorContent"))
			
	If StringNotEmptyOrNull(lngAboutId) Then
		blnSuccess = UpdateAboutText(lngAboutId,strAboutText)		
	Else
		blnSuccess = AddAboutText(strAboutText,lngAboutId)
	End If
		
	If blnSuccess Then
		'PageRedirect(REPORT_PAGE)
		displayMessage = "The information was successfully saved to the database."
	Else
		displayMessage = "An error occurred while trying to save information to the database."
	End If
End If

'Else
'	lngAboutId = GetQryString("id")
'	If StringNotEmptyOrNull(lngAboutId) Then
'		Set oAbout = New cAbout
'		Set collAbout = New cAbout
'		collAbout.ID = lngAboutId
'		Call collAbout.GetAboutTextById()
'		For Each oAbout In collAbout.Abouts.Items		
'			lngAboutId = oAbout.ID
'			strAboutText = oAbout.AboutText
'			Set oAbout = Nothing
'		Next
'		Set collAbout = Nothing
'	End If

Set oAbout = New cAbout
Set collAbout = New cAbout
Call collAbout.GetAboutText()
For Each oAbout In collAbout.Abouts.Items		
	lngAboutId = oAbout.ID
	strAboutText = oAbout.AboutText			
	Set oAbout = Nothing
Next
Set collAbout = Nothing

'End If
%>
<HTML>
<HEAD>
<TITLE><%=PAGE_TITLE%></TITLE>
<script src="/cms/lib/__dom.js" type="text/javascript"></script>
<script language="Javascript1.2" type="text/javascript">
function checkForm() {	
    var domEditor = findDOM('editorContent');
	domEditor.value = oEdit1.getHTML();
	var frm = document.forms['form1'];
	var mode = frm.hidMode;
	mode.value = 'save';
	frm.submit();
}
</script>
<!--#include virtual="/cms/styles.asp"-->
<script language=JavaScript src='/cms/scripts/innovaeditor.js'></script>
</HEAD>
<BODY>
<form name="form1" id="form1" method="post" action="<%=EDIT_PAGE%>">
<input type="hidden" id="hidAboutId" name="hidAboutId" value="<%=lngAboutId%>">
<input type="hidden" name="hidMode" id="hidMode" value="<%=strMode%>">
<input type="hidden" name="editorContent" id="editorContent" value="">
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" style="border-bottom-style:solid;border-width:1px;border-color:#666666;">
	<!--#include virtual="/cms/common/__header.asp"-->
	<!--#include virtual="/cms/common/__titlebar.asp" -->
	<!--include virtual="/cms/common/__begin_bodywrap.asp"-->
	<tr>
		<td width="100%">
			<table width="100%" border="0" cellpadding="0" cellspacing="0" height="100%">
				<tr>
					<td id="leftnav">
						<!-- #include virtual="/cms/sidenav.asp" -->
					</td>
					<td id="mainbody">		
						<!-- START PAGE BODY TOOLBAR -->	
						<div style="border-bottom:1px solid <%=abelard_border_color%>;width:100%;padding-bottom:20px;">
							<div style="align:left;width:<%=CMS_PAGE_WIDTH%>;">	
								<span style="float:right;padding-right:10px;">
									<A href="#" onClick="return checkForm();" class="menu" title="Save Information">
										<IMG height="16" alt="Save Information" src="<%=CMS_IMAGE_PATH%>/saveitem.gif" width="16" hspace="5" align="absmiddle">Save Information
									</A>
								</span>
							</div>
						</div>
						<!-- END PAGE BODY TOOLBAR -->
						<table border="0" cellpadding="0" cellspacing="0" width="<%=CMS_PAGE_WIDTH%>">
							<tr>
								<td width="50" valign="top"><br><img src="<%=CMS_IMAGE_PATH %>/<%=PAGE_IMAGE%>"></td>
								<td style="width:auto;">
									<p class="admin-instruction"><% If IS_EDIT_PAGE Then echo(EDIT_INSTRUCTIONS) Else echo(REPORT_INSTRUCTIONS) %></p>
									<p style="color:red;"><%If StringNotEmptyOrNull(displayMessage) Then echo("<br>" & displayMessage)%></p>	
									<table class="table-container" width="<%=CMS_PAGE_WIDTH%>">					
										<tr>
											<th align="left" width="200">About Text</th>
										</tr>
										<tr>
											<td>
	                                            <textarea id="taAboutText" name="taAboutText" rows="24" cols="80" class="admin-section"><%=strAboutText%></textarea>
	                                            <script>
		                                            var oEdit1 = new InnovaEditor("oEdit1");
		                                            oEdit1.REPLACE("taAboutText");
	                                            </script>												
											</td>
										</tr>
									</table>
	<!--#include virtual="/cms/common/__end_bodywrap.asp"-->
</table>
</form>
<!--#include virtual="/cms/common/__footer.asp"-->

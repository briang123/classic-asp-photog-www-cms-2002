<% 
Option Explicit 
Response.Buffer = True
'On Error Resume Next
%>
<!-- #include file="__globals.asp" -->
<!-- #include file="__dbfunct.asp" -->
<!-- #include file="__common.asp" -->
<!-- #include file="__library.asp" -->
<!-- #include file="__security.asp" -->
<!-- #include file="__debug.asp" -->
<html>
<head>
<title><%=COMPANY_NAME%> - Online Invoice Tracker</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="Javascript1.2" type="text/javascript">
<!-- 
function RequestSearch(frm) {
	frm.action="requests.asp?page=0&optionId=" + document.frmSearch.selSearchType.selectedIndex + "&search=" + escape(document.frmSearch.txtSearch.value) + ";"
	frm.submit();
}

function popup(strPageToOpen,win,h,w) {
	var winl = (screen.width - w) / 2;
	var wint = (screen.height - h) / 2;	
	popupWin= window.open(strPageToOpen,win,"height="+h+",width="+w+",left="+wint+",top="+winl+",status=no,toolbar=no,menubar=no,location=no,scrollbars=no,resizable=no");
}
//-->
</script>
<script language="JavaScript1.2" src="<%=strIncludePath%>/__dom.js" type="text/javascript"></script>
<style type="text/css">
<!--
body,td,th,.error {
	font-family: Arial, Helvetica, Geneva, sans-serif;
	font-size: 12px;
	font-weight: Normal;
}
body {
    MARGIN-LEFT: 0px;
    MARGIN-TOP: 0px;
}
th {
	font-size: 13px;
	font-weight: Bold;
}
th.fontWhite {
	color:#FFFFFF;
}
th.fontBlack {
	color:#000000;
}
.error {
	color: #FF0000;
	font-weight: Bold;
}
.LeftNavHeaderWhite {
	color:#FFFFFF;
	font-weight:bold;
	font-size:12px;
	font-family:Arial,Helvetica,Geneva,sans-serif
}
.LeftNavHeaderBlack {
	color:#000000;
	font-weight:bold;
	font-size:12px;
	font-family:Arial,Helvetica,Geneva,sans-serif
}

DIV.LoginError {
    BORDER-RIGHT: #808080 1px solid;
    PADDING-RIGHT: 15px;
    BORDER-TOP: #808080 1px solid;
    PADDING-LEFT: 15px;
    FONT-SIZE: 12px;
    PADDING-BOTTOM: 5px;
    BORDER-LEFT: #808080 1px solid;
		HEIGHT: 200px;
    WIDTH: 400px;
    COLOR: #FF0000;
    PADDING-TOP: 5px;
    BORDER-BOTTOM: #808080 1px solid;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    BACKGROUND-COLOR: beige
}
DIV.LoginError2 {
    BORDER-RIGHT: #808080 1px solid;
    PADDING-RIGHT: 20px;
    BORDER-TOP: #808080 1px solid;
    PADDING-LEFT: 20px;
    FONT-SIZE: 12px;
    PADDING-BOTTOM: 10px;
    BORDER-LEFT: #808080 1px solid;
		HEIGHT: 200px;
    WIDTH: 400px;
    COLOR: #FFFFFF;
    PADDING-TOP: 15px;
    BORDER-BOTTOM: #808080 1px solid;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    BACKGROUND-COLOR: #4A6999
}
DIV.loading{
	BACKGROUND-COLOR: #FFFFFF
  Z-INDEX: 100;
  LEFT: 43%;
  FLOAT: none;
  POSITION: absolute;
  TOP: 30%;
	COLOR: #666666;
}
div.login
{
    BORDER-RIGHT: #808080 1px solid;
    PADDING-RIGHT: 2px;
    BORDER-TOP: #808080 1px solid;
    PADDING-LEFT: 2px;
    FONT-SIZE: 12px;
    PADDING-BOTTOM: 2px;
    BORDER-LEFT: #808080 1px solid;
    WIDTH: 400px;
    COLOR: #000000;
    PADDING-TOP: 2px;
    BORDER-BOTTOM: #808080 1px solid;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    BACKGROUND-COLOR: beige
}
div.login2
{
    BORDER-RIGHT: #808080 1px solid;
    PADDING-RIGHT: 2px;
    BORDER-TOP: #808080 1px solid;
    PADDING-LEFT: 2px;
    FONT-SIZE: 12px;
    PADDING-BOTTOM: 2px;
    BORDER-LEFT: #808080 1px solid;
    WIDTH: 400px;
    COLOR: #FFFFFF;
    PADDING-TOP: 2px;
    BORDER-BOTTOM: #808080 1px solid;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    BACKGROUND-COLOR: #4A6999
}
DIV.box
{
    BORDER-RIGHT: #808080 1px solid;
    PADDING-RIGHT: 2px;
    BORDER-TOP: #808080 1px solid;
    PADDING-LEFT: 2px;
    FONT-SIZE: 12px;
    PADDING-BOTTOM: 2px;
    BORDER-LEFT: #808080 1px solid;
    WIDTH: 740px;
    COLOR: #000000;
    PADDING-TOP: 2px;
    BORDER-BOTTOM: #808080 1px solid;
    FONT-FAMILY: Arial, Helvetica, Sans-serif;
    BACKGROUND-COLOR: #f5f5dc
}
.tblfiles {
	FONT-SIZE: xx-small;
	FONT-FAMILY: Tahoma;
}
.inpfiles {
	font:8pt verdana,arial,sans-serif;
}
.selfiles {
	height: 22px; 
	top:2;
	font:8pt verdana,arial,sans-serif;
}	
.bar {
	BORDER: #666666 1px solid; 
	BACKGROUND: #FF8800; 
	WIDTH: 100%; 
	HEIGHT: 20px;
}
.imgbar {
	BORDER-BOTTOM: #666666 1px solid;
	BACKGROUND: #FF8800; 
	WIDTH: 100%; 
	HEIGHT: 19px;
	TEXT-ALIGN: left;
}	
.imgList {
	BORDER-LEFT: #666666 1px solid;
	BORDER-RIGHT: #666666 1px solid;
	BORDER-TOP: #666666 1px solid;
	BORDER-BOTTOM: #666666 1px solid;
	WIDTH: 210px;
}
.message {
	COLOR: #666666;
	FONT-WEIGHT: BOLD;
	FONT-SIZE: 10px;
	FONT-FAMILY: Arial, Helvetica, Geneva, Sans-serif;
}
.header {
	COLOR: #FF8800;
	FONT-WEIGHT: BOLD;
	FONT-SIZE: 14px;
	FONT-FAMILY: Arial, Helvetica, Geneva, Sans-serif;
}
.imgText {
	COLOR: #666666;
	FONT-SIZE: 10px;
	FONT-FAMILY: Arial, Helvetica, Geneva, Sans-serif;
}
.msg {
	COLOR: #FF0000;
	FONT-WEIGHT: BOLD;
}
.toolBar {
	HEIGHT:300px;
	WIDTH:160px;
	BORDER-LEFT: #666666 1px solid;
	BORDER-RIGHT: #666666 1px solid;
	BORDER-BOTTOM: #666666 1px solid;
}
iframe {
	margin-left: 0px;
	margin-top: 0px;

}
	a         { color: #0000BB; text-decoration: none; }
	a:hover   { color: #FF0000; text-decoration: underline; }
	.headline { font-family: arial black, arial; font-size: 28px; letter-spacing: -1px; }
	.headline2{ font-family: verdana, arial; font-size: 12px; }
	.subhead  { font-family: arial, arial; font-size: 18px; font-weight: bold; font-style: italic; }
	.backtotop     { font-family: arial, arial; font-size: xx-small;  }
	.code     { background-color: #EEEEEE; font-family: Courier New; font-size: x-small;
							margin: 5px 0px 5px 0px; padding: 5px;
							border: black 1px dotted;
						}
/*	font { font-family: arial black, arial; font-size: 28px; letter-spacing: -1px; }*/
	
	
.menuHead  {
	color: black;
	font-weight: bold;
	font-size: 11px;
	font-family: Arial, Helvetica, Geneva, sans-serif;
	text-decoration: none }
.SubMenuIndent {
	color: black;
	font-size: 12px;
	font-weight: bold;
	font-family: Arial, Helvetica, Geneva, sans-serif;
	padding-left: 20px; }
.nameIndent {
	color: black;
	font-size: 11px;
	font-family: Arial, Helvetica, Geneva, sans-serif;
	padding-left: 10px; }
.required {
	color: #FF8800;
	font-size: 9px;
}
.poweredby {
	color: #666666;
	font-size: 9px;
}
//-->
</style>
</head>
<% Call HTMLComment("HEADER SECTION",1) %>
<body bgcolor="<%=strBodyColor%>">
<table border="<%=intBorder%>" width="990px" height="500px" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
	<tr valign="top">
		<td valign="top" height="100%">
			<table width="990px" border="<%=intBorder%>" cellpadding="0" cellspacing="0" bgcolor="<%=strHeadColor%>">
				<tr valign="top">
					<td width="990px" valign="top">
						<table width="100%" border="<%=intBorder%>" cellpadding="0" cellspacing="0">
							<tr>
								<td align="left" width="210px">
									<img src="<%=strImagePath%>/complogos/<%=GetAppVariable("__CompanyLogo")%>" width="210px" height="55px" border="0" alt="<%=GetAppVariable("__CompanyName")%>">							
								</td>
								<td align="right" valign="bottom">
									<table cellpadding="0" cellspacing="0" width="100%" border="0" align="right">
										<tr>
											<td><img src="<%=strImagePath%>/1x1.gif" width="575px" height="1px"></td>
											<td align="left"><b>Keyword Search:</b></td>
										</tr>
										<form name="frmSearch" method="post" action="./requests.asp?page=0">
										<tr align="right">
											<td align="right" valign="bottom">
											<!--
												<a href="javascript:alert('This will be online help popup window.');"><img src="<%=strImagePath%>/help.gif" border="0"></a>&nbsp;&nbsp;
												<a href="javascript:alert('This will link back to support application homepage');"><img src="<%=strImagePath%>/home.gif" border="0"></a>&nbsp;&nbsp;
											-->
												<a href="javascript:popup('<%=strIncludePath%>/popups/profile.asp','winProfile',400,400);"><img src="<%=strImagePath%>/profilestar.gif" border="0"></a>&nbsp;&nbsp;
											<!--<img src="<%=strImagePath%>/login.gif" border="0">&nbsp;&nbsp;-->
											</td>										
											<td>
												<select name="selSearchType">
													<option value="0">Select One</option>
													<option value="1" <% if GetFormPost("selSearchType") = 1 Then response.write "selected"%>>Request ID</option>
													<option value="2" <% if GetFormPost("selSearchType") = 2 Then response.write "selected"%>>Title</option>
													<option value="3" <% if GetFormPost("selSearchType") = 3 Then response.write "selected"%>>Batch Date</option>
													<option value="4" <% if GetFormPost("selSearchType") = 4 Then response.write "selected"%>>Open Date</option>
													<option value="5" <% if GetFormPost("selSearchType") = 5 Then response.write "selected"%>>Close Date</option>
												</select>
												<input type="text" name="txtSearch" size="10" value="<%=GetFormPost("txtSearch")%>">
												<input type="image" name="butSearch" src="<%=strImagePath%>/goButton.gif" alt="Search">
											</td>
										</tr>
										</form>								
									</table>
								</td>								
							</tr>
						</table>											
					</td>
				</tr>
				<tr>
					<td width="100%" valign="top">
						<table border="<%=intBorder%>" cellpadding="0" cellspacing="0" width="100%">
							<tr valign="top">
								<td rowspan="100%" background="<%=strImagePath%>/rightnavborder.gif" width="10px" valign="top">&nbsp;</td>
								<td width="980px" valign="top">
									<table width="980px" border="<%=intBorder%>" cellspacing="0" cellpadding="0">
										<tr valign="top">
											<td width="980px" valign="top">
												<table width="980px" border="<%=intBorder%>" cellpadding="0" cellspacing="0">
													<tr>
														<td width="980px" height="10px" background="<%=strImagePath & strTopDividerBarImage%>">
															<img src="<%=strImagePath%>/1x1.gif" width="1px" height="10px">
														</td>
													</tr>
												</table>
											</td>
										</tr>
										<tr valign="top">
<% Call HTMLComment("HEADER SECTION",2) %>										
											<td rowspan="100%" width="100%" valign="top">
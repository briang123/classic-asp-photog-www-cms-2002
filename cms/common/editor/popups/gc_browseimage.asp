<%
'Option Explicit
Response.Buffer = False
%>
<!-- #include File="../../gc_common.asp" -->
<!-- #include file="../../gc_fsoConfig.asp" -->
<!-- #include file="../../gc_fsoCommon.asp" -->
<%
' We need to determine what operation we are performing
If Request.QueryString("action") = 2 And Request.QueryString("file") <> "" Then
	Call DeleteFSOFile(Request.QueryString("file"))
End If
%>

<html>
<head>
<script language="JavaScript1.2" type="text/javascript">
<!--
// Selects the current table row when it is clicked on
function selectRow(row){
	var nRows;
	var table;

	table = document.getElementById("fileTable");
	nRows = table.rows.length;
	for (var iCounter=0; iCounter<nRows; iCounter++){
		table.rows(iCounter).style.backgroundColor="#ffffff";
	}
	row.style.backgroundColor="#8DBCEB";
}

// Gives the end user a loading message during processing
// This function is called from the body onLoad event.
function remLoadMessage(){
	document.getElementById("loadMessage").style.display = "none";
}
//-->
</script>
<link href="<%=Application("SupportAppIncludePath")%>/gc_imgMgmt.css" type="text/css" rel="stylesheet" media="screen">
</head>
<body bgcolor="white" topmargin="0" leftmargin="0" onLoad="remLoadMessage();" link="#0000FF" vlink="#0066FF" alink="#0066FF">
	<div id="loadMessage" class="loadiframe" align="center">
		<h3><font face="Tahoma">loading...</font></h3>
		<img src="<%=Application("SupportAppImagePath")%>/progressbar.gif" border="0" width="82px" height="10px">
	</div>
	<table border=0 cellpadding=0 cellspacing=0 width=360>
		<tr>
			<td>
				<div class="bar" style="padding-left: 5px; border-left:0px; border-right:0px;">
					<font size="2" face="tahoma" color="white"><img border="0" src="<%=strHeaderImage%>">&nbsp;<b>File Name</b></font>
				</div>
			</td>
		</tr>
	</table>
	<div style="overflow:auto;height:290;width:360;border:0px;">
		<table id="fileTable" border="0" cellpadding="3" cellspacing="0" width="350">
		<% 
		Dim fsoSingleFolder
		fsoSingleFolder = request.QueryString("dir")
		If fsoSingleFolder = "" Then
			fsoSingleFolder = Left(strCeilingFolder,Len(strCeilingFolder)-1)
		End If
		Set aFolder = objFSO.GetFolder(fsoSingleFolder & "\")

		dim strAction
		If request.QueryString("action") = 1 Then
			strAction = "SELECT"
		ElseIf request.QueryString("action") = 2 Then
			strAction = "DELETE"
		End If
		Call GetFSOFiles2(aFolder,strFileExtFilter,strAction)
		%>
		</table>
	</div>
</body>
</html>

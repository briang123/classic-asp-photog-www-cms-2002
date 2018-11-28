<!-- #include file="../../../gc_common.asp" -->
<!-- #include file="../../../gc_fsoConfig.asp" -->
<!-- #include file="../../../gc_fsoCommon.asp" -->
<html>
<head>
	<title>File Uploader</title>
<script language="JavaScript1.2" type="text/javascript">
<!--
// Adds a new file attachment component dynamically on the screen
var nfiles = 1;
function Expand() {
	nfiles++
	if (nfiles > 5) {
		alert('You can only upload a maximum of 5 files at a time.');
		return false;
	} else {
	  var adh = 'File Attachment '+nfiles+':&nbsp;<input type="file" size="40" name="file'+nfiles+'" style="font:8pt verdana,arial,sans-serif"><br/>';
  	files.insertAdjacentHTML('BeforeEnd',adh);
	  return false;
	}
}
//-->
</script>
<link href="<%=Application("SupportAppIncludePath")%>/gc_imgMgmt.css" type="text/css" rel="stylesheet" media="screen">
</head>
<body onLoad="setParent();" bgcolor="white" topmargin="0" leftmargin="0" link="#0000FF" vlink="#0066FF" alink="#0066FF">
<br/>
<table border="0" cellpadding="1" cellspacing="0" bgcolor="#666666" width="99%">
	<tr>
		<td rowspan="100%" width="5px" bgcolor="#FFFFFF"><img src="<%=Application("SupportAppImagePath")%>/1x1.gif" width="5px" height="1px" border="0"></td>
		<td>
			<table width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="#FFFFFF">
				<tr>
					<td colspan="100%">
					<div class="imgbar" style="BORDER-BOTTOM: #666666 1px solid;padding-left:5px;width=100%">
						<font size="2" face="tahoma" color="white"><img border="0" src="<%=strHeaderImage%>">&nbsp;<b>File Upload:</b></font>
					</div>
					</td>
				</tr>
				<tr><td colspan="100%" align="center">
				<% select case request.QueryString("msg")
						case "exist"
							response.write "One of the uploaded files already exist.<br/>"
						case "fail"
							response.write "The upload process failed.</br>"
						case "success"
							response.write "The upload process was successful.<br/>"
					end select
				%></td></tr>
				<tr><td rowspan="100%" width="5px" bgcolor="#FFFFFF"><img src="<%=Application("SupportAppImagePath")%>/1x1.gif" width="5px" height="1px" border="0"></td></tr>
				<form name="formUpload" id="formUpload" method="post" enctype="multipart/form-data" action="gc_fileUploadProcessing.asp">
				<tr>
					<td>Select Upload Folder:&nbsp;
						<SELECT ID="selFolderList" NAME="selFolderList">
						<% Call BuildFSOFolderCombo(objFolder,"") %>
						</SELECT>
					</td>	
				</tr>
				<tr><td height="5px"><img src="<%=Application("SupportAppImagePath")%>/1x1.gif" width="1px" height="5px" border="0"></td></tr>
				<tr>
					<td>
						<div id="files">
							File Attachment 1:&nbsp;<input type="file" size="40" name="file1" style="font:8pt verdana,arial,sans-serif"><br/>
						</div>
						<br/>
					</td>
				</tr>
				<tr>
					<td align="center">
						<input type="button" value="Add Attachment" OnClick="return Expand();" style="height: 22px;font:8pt verdana,arial,sans-serif">&nbsp;&nbsp;
						<% if Session("Email") <> "bgaines@gainesconsulting.com" Then 
								Response.write "<b>No Permissions to Upload File now</b>"
							else %>
						<input type="submit" value="Upload Files" style="height: 22px;font:8pt verdana,arial,sans-serif">
						<% end if %>
						&nbsp;&nbsp;
						<input type="button" value="Cancel" onClick="window.close();" style="height: 22px;font:8pt verdana,arial,sans-serif">&nbsp;&nbsp;
					</td>
					</form> 
				</tr>
			</table>			
		</td>
	</tr>
</table>
</body>
</html>
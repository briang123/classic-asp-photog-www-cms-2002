<% Response.Buffer = True %>
<!-- #include File="../../gc_common.asp" -->
<!-- #include file="../../gc_fsoConfig.asp" -->
<!-- #include file="../../gc_fsoCommon.asp" -->
<html>
<head>
<title>Image Management Console</title>
<SCRIPT language="JavaScript1.2" type="text/javascript" defer>
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

function _CloseOnEsc() {
  if (event.keyCode == 27) { window.close(); return; }
}

function _getTextRange(elm) {
  var r = elm.parentTextEdit.createTextRange();
  r.moveToElementText(elm);
  return r;
}

var fImageLoaded
window.onerror = HandleError

// Handle javascript errors
function HandleError(message, url, line) {
  var str = "An error has occurred in this dialog." + "\n\n"
  + "Error: " + line + "\n" + message;
  alert(str);
	window.close();
  return true;
}

// Initialize the values when page loads
function Init() {

  var elmSelectedImage;
  var htmlSelectionControl = "Control";
  var globalDoc = window.dialogArguments;
  var grngMaster = globalDoc.selection.createRange();
  	
  // event handlers  
  document.body.onkeypress = _CloseOnEsc;
  btnOK.onclick = new Function("btnOKClick()");

	fImageLoaded = false;
  if (globalDoc.selection.type == htmlSelectionControl) {
    if (grngMaster.length == 1) {
      elmSelectedImage = grngMaster.item(0);
      if (elmSelectedImage.tagName == "IMG") {
				fImageLoaded = true;
        if (elmSelectedImage.src) {

//					inpImgUrl.value = elmSelectedImage.src;

					// display image file in previewer
					showPreview(elmSelectedImage.src,elmSelectedImage.width,elmSelectedImage.height,"");
					
          inpImgUrl.value 					= elmSelectedImage.src.replace(/^[^*]*(\*\*\*)/, "$1");  // fix placeholder src values that editor converted to abs paths
					inpImgWidth.value  				= elmSelectedImage.width;
					inpImgHeight.value  			= elmSelectedImage.height;
          inpVSpace.value         	= elmSelectedImage.vspace;
          inpHSpace.value       		= elmSelectedImage.hspace;
          inpImgBorder.value        = elmSelectedImage.border;
          inpImgAlt.value          	= elmSelectedImage.alt;
          inpImgAlign.value        	= elmSelectedImage.align;
        }
      }
    }
  }
	inpImgUrl.focus();
}

// validate valid number
function _isValidNumber(txtBox) {
  var val = parseInt(txtBox);
  if (isNaN(val) || val < 0 || val > 999) { return false; }
  return true;
}

// Perform operations once user clicks the "Insert Image" button
function btnOKClick() {
  var elmImage;
  var intAlignment;
  var htmlSelectionControl = "Control";
  var globalDoc = window.dialogArguments;
  var grngMaster = globalDoc.selection.createRange();
  
  // error checking
  if (inpHSpace.value && !_isValidNumber(inpHSpace.value)) {
    alert("Horizontal spacing must be a number between 0 and 999.");
    inpHSpace.focus();
    return;
  }
  if (inpVSpace.value && !_isValidNumber(inpVSpace.value)) {
    alert("Vertical spacing must be a number between 0 and 999.");
    inpVSpace.focus();
    return;
  }

  // delete selected content and replace with image
  if (globalDoc.selection.type == htmlSelectionControl && !fImageLoaded) {
    grngMaster.execCommand('Delete');
    grngMaster = globalDoc.selection.createRange();
  }
    
	// new image creation ID
  idstr = "\" id=\"556e697175657e537472696e67";     
	if (!fImageLoaded) {
    grngMaster.execCommand("InsertImage", false, idstr);
    elmImage = globalDoc.all['556e697175657e537472696e67'];
    elmImage.removeAttribute("id");
    elmImage.removeAttribute("src");
    grngMaster.moveStart("character", -1);
  } else {
    elmImage = grngMaster.item(0);
    if (elmImage.src != inpImgUrl.value) {
      grngMaster.execCommand('Delete');
      grngMaster = globalDoc.selection.createRange();
      grngMaster.execCommand("InsertImage", false, idstr);
      elmImage = globalDoc.all['556e697175657e537472696e67'];
      elmImage.removeAttribute("id");
      elmImage.removeAttribute("src");
      grngMaster.moveStart("character", -1);
			fImageLoaded = false;
    }
    grngMaster = _getTextRange(elmImage);
  }

	if (fImageLoaded) {
    elmImage.style.width = inpImgWidth.value;
    elmImage.style.height = inpImgHeight.value;
  }

  if (inpImgUrl.value.length > 2040) {
  	inpImgUrl.value = inpImgUrl.value.substring(0,2040);
	}
  
	elmImage.src = inpImgUrl.value

  if (inpHSpace.value != "") { elmImage.hspace = parseInt(inpHSpace.value); }
  else { elmImage.hspace = 0; }

  if (inpVSpace.value != "") { elmImage.vspace = parseInt(inpVSpace.value); }
  else { elmImage.vspace = 0; }
  
  elmImage.alt = inpImgAlt.value;

  if (inpImgBorder.value != "") { elmImage.border = parseInt(inpImgBorder.value); }
  else { elmImage.border = 0; }

  elmImage.align = inpImgAlign.value;
  grngMaster.collapse(false);
  grngMaster.select();
  window.close();
}

// Set values based on the image that is selected in the IFrame window
var imgSource
function selectImage(sURL,w,h,imgname) {

	// set up the file parameters when it's selected.
	imgSource							=	'<%=Application("SupportAppHostName")%>' + sURL;
	inpImgUrl.value				=	imgSource;
	inpImgWidth.value 		= w;
	inpImgHeight.value 		= h;
	inpImgAlt.value 			= "";
	inpVSpace.value				= 0;
	inpHSpace.value				= 0;
	inpImgBorder.value		= 0;
	inpImgAlign.value			= "";

	// preview the image file
	showPreview(imgSource);
}

function deleteImage(sURL) {
	if (confirm("Are you sure would like to delete this file?") == true) {
		reloadFileBrowser(sURL);
	}
}

// Preview the image
function showPreview(sURL) {
	
	// update the preview window
	divImg.style.visibility = "visible"
	divImg.innerHTML = "<img id='PREVIEWPIC' vspace='30' src='" + sURL + "'>";

	var width = PREVIEWPIC.width;
	var height = PREVIEWPIC.height;
	var resizedWidth = 180; //150
	var resizedHeight = 200; //170
	var Ratio1 = resizedWidth/resizedHeight;
	var Ratio2 = width/height;

	// preview the image in different sizes
	if(Ratio2 > Ratio1) {
		if(width*1>resizedWidth*1)
			PREVIEWPIC.width=resizedWidth;
		else
			PREVIEWPIC.width = (width>180) ? 180 : width;
	} else {
		if(height*1>resizedHeight*1)
			PREVIEWPIC.height=resizedHeight;
		else
			PREVIEWPIC.height = (height>200) ? 200 : height;
	}

	// preview the image in the previewer
	divImg.style.visibility = "visible"
	
	Ratio1=0;
	Ratio2=0;
	
}

// Reloads the IFramed file list with a new set of files based on the directory path
function reloadFileBrowser(sURL) {
	document.all.IMGPICK.src="gc_browseimage.asp?dir="+document.frmFileMgmt.selFolderList.value+'&action=<%=request.QueryString("action")%>&file='+escape(sURL);
}					
//-->
</script>
<link href="<%=Application("SupportAppIncludePath")%>/gc_imgMgmt.css" type="text/css" rel="stylesheet" media="screen">
</head>
<body onLoad="Init();" bgcolor="#FFFFFF">
<table width="100%" cellpadding="0" cellspacing="0" border="0">
	<tr>
		<td height="10px" colspan="100%" bgcolor="#666666"><img src="<%=Application("SupportAppImagePath")%>/1x1.gif" border="0" width="1px" height="10px"></td>
	</tr>
	<tr>
		<td>
			<table border="0" cellpadding="5" cellspacing="0">
				<tr>
					<td>
						<table border="0" cellpadding="3" cellspacing="3" align="left">
							<tr>
								<td valign=top rowspan="2">
									<table border=0 height=48 cellpadding=0 cellspacing=0>
										<tr>
											<!-- START IMAGE FOLDER COMBO BOX GENERATION -->
											<form method="post" id="frmFileMgmt" name="frmFileMgmt" action="">
											<td><img border="0" src="<%=strHeaderImage%>">&nbsp;<b>Select folder:</b></td>
											<td colspan="2">
												<SELECT ID="selFolderList" NAME="selFolderList" onClick="reloadFileBrowser('');">
												<% Call BuildFSOFolderCombo(objFolder,Request.Form("selFolderList")) %>
												</SELECT>											
											</td>
											</form>
											<!-- END IMAGE FOLDER COMBO BOX GENERATION -->											
										</tr>
									</table>			
									<table cellpadding="0" cellspacing="0" border="0">
										<tr>
											<td valign="top">
												<!-- START FILE LIST WINDOW -->
												<div style="BORDER-LEFT: #666666 1px solid;BORDER-RIGHT: #666666 1px solid;BORDER-BOTTOM: #666666 1px solid;">								
													<iframe name="IMGPICK" src="gc_browseimage.asp?action=<%=request.QueryString("action")%>" height="290px" width="360px" frameborder="0" scrolling="no"></iframe>								
												</div>
												<!-- END FILE LIST WINDOW -->												
											</td>
											<td width="5px">
												<!-- SPACER -->
												<img src="<%=Application("SupportAppImagePath")%>/1x1.gif" border="0" width="5px" height="1px">
											</td>
											<td valign="top" align="center" class="imgList" bgcolor="#FFFFFF" height="100%">
												<div class="imgbar" style="width:100%;padding-left:5px;">
													<font size="2" face="tahoma" color="white"><img border="0" src="<%=strHeaderImage%>">&nbsp;<b>Image Preview</b></font>
												</div>
												<div id="divImg" style="width:180;height:200;">
													<IMG ID="PREVIEWPIC" NAME="PREVIEWPIC" bgcolor="#ffffff" src=<%=Application("SupportAppIncludePath")%>/editor/imagePreview.gif alt="Preview" align="absmiddle" valign="middle"> 					
												</div>
												<div id="divMsg" class="msg"><p>
												<%
													if Request.QueryString("file") <> "" then
														dim arrDelFile,intUBound
														arrDelFile = Split(Request.QueryString("file"),"/")
														intUBound = UBound(arrDelFile)
														Response.Write(arrDelFile(intUBound) & " has been deleted.")
													end if
												%>
												</p></div>
											</td>
										</tr>									
									</table>	
									<table border=0 width=340 cellpadding=0 cellspacing=1>
										<tr>
											<td>Filename:</td>
											<td colspan=3><INPUT TYPE="text" size="40" NAME="inpImgUrl" value="" onChange="showPreview()" onFocus="select();">
											</td>		
										</tr>		
										<tr>
											<td>Alignment:</td>
											<td>
												<select ID="inpImgAlign" NAME="inpImgAlign">
													<option value="" selected>&lt;Not Set&gt;</option>
													<option value="absBottom">absBottom</option>
													<option value="absMiddle">absMiddle</option>
													<option value="baseline">baseline</option>
													<option value="bottom">bottom</option>
													<option value="left">left</option>
													<option value="middle">middle</option>
													<option value="right">right</option>
													<option value="textTop">textTop</option>
													<option value="top">top</option>						
												</select>
											</td>
											<td>Image border:</td>
											<td>
												<select id="inpImgBorder" name="inpImgBorder">
													<option value=0>0</option>
													<option value=1>1</option>
													<option value=2>2</option>
													<option value=3>3</option>
													<option value=4>4</option>
													<option value=5>5</option>
												</select>
											</td>					
										</tr>
										<tr>
											<td>ALT Text:</td>
											<td colspan=3><INPUT type="text" id="inpImgAlt" name="inpImgAlt" size="39" onFocus="select();"></td>		
										</tr>							
										<tr>
											<td>Image Width:</td>
											<td><INPUT type="text" ID="inpImgWidth" NAME="inpImgWidth" size=2 onFocus="select();"></td>
											<td>Horizontal Spacing :</td>
											<td><INPUT type="text" ID="inpHSpace" NAME="inpHSpace" size=2 onFocus="select();">
											</td>
										</tr>				
										<tr>
											<td>Image Height:</td>
											<td><INPUT type="text" ID="inpImgHeight" NAME="inpImgHeight" size=2 onFocus="select();"></td>
											<td>Vertical Spacing :</td>
											<td><INPUT type="text" ID="inpVSpace" NAME="inpVSpace" size=2 onFocus="select();">
											</td>
										</tr>
									</table>
								</td>	
							</tr>		
						</table>						
					</td>			
				</tr>
				<tr>
					<td align=center colspan=2>
						<table cellpadding=0 cellspacing=0 align=center>
							<tr>
								<td>
									<button id=btnOK type=submit tabIndex=40 style="height: 22px;font:8pt verdana,arial,sans-serif">Insert Image</button>&nbsp;&nbsp;
									<button id=btnCancel type=reset tabIndex=45 style="height: 22px;font:8pt verdana,arial,sans-serif" onClick="window.close();">Cancel</button>
								</td>
							</tr>
						</table>
					</td>
				</tr>	
			</table>
		
		</td>
	</tr>
</table>
<input type=text style="display:none;" id="inpActiveEditor" name="inpActiveEditor" contentEditable="true">
<%
Set objFSO = Nothing
Set objFolder = Nothing
%>
</body>
</html>
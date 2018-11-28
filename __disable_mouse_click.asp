<% If DISABLE_IMAGE_RIGHT_MOUSE_CLICK Then %>
<script type="text/javascript" language="javascript">
var clickmessage="Right click is disabled on <%=COMPANY_NAME%> images. Please contact <%=PHOTOGRAPHER_FNAME%> at \n<%=COMPANY_PHONE%> if you would like to learn more or schedule a session."
function disableclick(e) {
	if (document.all) {
		if (event.button==2||event.button==3) {
			if (event.srcElement.tagName=="IMG"){
				alert(clickmessage);
				return false;
			}
		}
	} else if (document.layers) {
		if (e.which == 3) {
			alert(clickmessage);
			return false;
		}
	} else if (document.getElementById){
		if (e.which==3&&e.target.tagName=="IMG"){
			alert(clickmessage)
			return false
		}
	}
}

function associateimages(){
	for(i=0;i<document.images.length;i++) document.images[i].onmousedown=disableclick;
}
if (document.all)
	document.onmousedown=disableclick
else if (document.getElementById)
	document.onmouseup=disableclick
else if (document.layers)
	associateimages()
</script>

<% 
	If err <> 0 Then
		PageRedirect("error.asp")
	End If

End If %> 
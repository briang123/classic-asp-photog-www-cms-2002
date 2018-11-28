<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/cms/common/__fsoConfig.asp" -->
<!-- #include virtual="/cms/common/__fsoCommon.asp" -->
<!-- #include virtual="/objects/cPhotos.asp" -->
<html>
<head>
<title>julie stark children photography in New York</title>
<title>julie stark children photography in New York</title>
<!-- #include file="__meta.asp" -->
<!-- #include file="__css.asp" -->
<%
Dim BODY_TITLE
BODY_TITLE = "home"

If RUN_HOME_PAGE_SLIDE_SHOW Then %>
	<script type="text/javascript" language="javascript">
	<!--
	var slidespeed = <%=SLIDE_SHOW_SLIDE_SPEED%>;
	var slideimages = new Array();
	<%
	Set oPhoto = New cPhotos
	Set collPhotos = New cPhotos
	Call collPhotos.GetSitePhotos()
	Dim imgArrCounter
	Dim startImage
	imgArrCounter = 0
	For Each oPhoto In collPhotos.Photos.Items 
		If oPhoto.WebPageId = WEB_PAGE_ID And oPhoto.ActiveFlag = True Then
		
		If imgArrCounter = 0 Then
			startImage = IMAGE_PATH & "/" & oPhoto.LargeImage
		End If	%> 
	slideimages[<%=imgArrCounter%>]='<%=IMAGE_PATH & "/" & oPhoto.LargeImage%>'; 
	<% imgArrCounter = imgArrCounter + 1
		End If
			Set oPhoto = Nothing
	Next
	Set collPhotos = Nothing
	%>
	var imageholder = new Array();
	var ie = document.all;
	for (i=0;i<slideimages.length;i++){
		imageholder[i] = new Image();
		imageholder[i].src = slideimages[i];
	}
	//-->
	</script>
<% End If %>
<script src="__menu.js" type="text/javascript"></script>
</head>
<body onload="preloadImages();">
<table width="900" height="562" border="0" cellpadding="0" cellspacing="0" align="center">
    <!-- #include file="__menu.asp" -->
        <td width="826" height="359" colspan="11" valign="top" align="center" style="background-color:<%=BodyContentBackColor%>;">
            <table width="826">
		        <tr>
		            <td><img src="images/spacer.gif" width="1" height="353" alt="" /></td>
		            <td>
                        <table width="810" style="border-top:1px solid black;border-right:1px solid black;border-bottom:1px solid black;border-left:1px solid black;">
                            <tr>
                            <% If RUN_HOME_PAGE_SLIDE_SHOW Then %>
	                            <td valign="middle" align="center">
		                            <img width="800" height="275" alt="<%=ALT_IMAGE_TEXT%>" src="<%=startImage%>" name="slide" border=0 style="filter:progid:DXImageTransform.Microsoft.Wipe(duration=<%=SLIDE_SHOW_DURATION%>,gradientsize=<%=SLIDE_SHOW_GRADIENT_SIZE%>, motion=<%=SLIDE_SHOW_MOTION%>)">
	                            </td>
                            <% Else 
	                            Set oPhoto = New cPhotos
	                            Set collPhotos = New cPhotos
	                            Call collPhotos.GetSitePhotos()
	                            For Each oPhoto In collPhotos.Photos.Items 
		                            If oPhoto.WebPageId = WEB_PAGE_ID And oPhoto.ActiveFlag = True Then	%>
	                            <td class="home-photo-display" valign="top" align="center" style="padding:5 0 5 0">
		                            <table height="275" cellpadding="0" cellspacing="0" border="0" class="home-photo-display-liner">
			                            <tr>
				                            <td valign="top"><img alt="<%=ALT_IMAGE_TEXT%>" src="<%=IMAGE_PATH%>/<%=oPhoto.LargeImage%>" height="275"></td>
			                            </tr>
		                            </table>
	                            </td>

	                            <%	End If
			                            Set oPhoto = Nothing
	                            Next
	                            Set collPhotos = Nothing
                            End If %>		
                            </tr>
                        </table>			        
                        <table width="810">
                            <tr>
                                <td align="right" style="padding:10 0 0 0;"><img src="images/home_tagline.jpg" /></td>
                            </tr>
                        </table>
                    </td>
		        </tr>
		    </table>	
        </td>
<!-- #include file="__partrow8_row9_10_copyright.asp" -->
</body>
</html>

<% 
If RUN_HOME_PAGE_SLIDE_SHOW And WEB_PAGE_ID = 2 Then %>
<script type="text/javascript" language="javascript"><!--	
var ie = document.all;
var whichimage=0;
var blenddelay=(ie) ? document.images.slide.filters[0].duration*1000 : 0;
function slideit() {
	if (!document.images) return
	if (ie) document.images.slide.filters[0].apply();
	document.images.slide.src = imageholder[whichimage].src;
	if (ie) document.images.slide.filters[0].play();
	whichimage =(whichimage<slideimages.length-1) ? whichimage+1 : 0;
	setTimeout("slideit()",slidespeed+500);
}
slideit();
//--></script>
<% End If %>
<!-- #include file="__disable_mouse_click.asp" -->

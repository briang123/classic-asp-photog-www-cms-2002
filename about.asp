<!-- #include virtual="/lib/__common.asp" -->
<!-- #include virtual="/lib/__globals.asp" -->
<!-- #include virtual="/objects/cAbout.asp" -->
<!-- #include virtual="/objects/cPhotos.asp" -->
<% 
Dim BODY_TITLE
BODY_TITLE = "About " & PHOTOGRAPHER_FNAME 
%>
<html>
<head>
<title>learn more about julie stark children photography in New York</title>
<title>learn more about julie stark children photography in New York</title>
<!-- #include file="__meta.asp" -->
<!-- #include file="__css.asp" -->
<script src="__menu.js" type="text/javascript"></script>
</head>
<body onload="preloadImages();">
<table width="900" height="562" border="0" cellpadding="0" cellspacing="0" align="center">
    <!-- #include file="__menu.asp" -->
        <td width="826" height="359" colspan="11" valign="top" align="center" style="background-color:<%=BodyContentBackColor%>;">
            <table width="826">
		        <tr>
		            <td><img src="images/spacer.gif" width="1" height="353" alt="" /></td>
                    <!--#include file="__sidephoto.asp"-->
		            <!--#include file="__bodytext.asp"-->
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

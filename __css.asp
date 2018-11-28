<%
Dim PHOTOGRAPHER_NAME
Dim BodyContentBackColor
Dim WebsiteContainerColor
Dim MenuColor
Dim OutsideColor
Dim HeaderColor
Dim MenuSelectedColor
Dim MenuBorderColor
Dim MenuSelectedFontColor
Dim MenuFontColor
Dim GalleryBackColor
Dim GalleryPhotoBorderColor
Dim BodyFontColor
Dim LinkFontColor
Dim PhotoDisplayBackColor
Dim FontFamily
Dim siteId
Dim ImagePath

siteId = 2
PHOTOGRAPHER_NAME = "Julie"
FontFamily = """century gothic"", sans-serif"
PhotoDisplayBackColor = "#F9F9E1" '"#FFFFFF"
LinkFontColor = "#88534F" '"#513F1B" '8E5252 
BodyFontColor = "#666666"
BodyContentBackColor = "#F5F5DD" '"#FFFFFF" 
'BodyContentBackColor = "#EEEEEE"
WebsiteContainerColor = "#DDDDDD"
MenuFontColor = "#FFFFFF"
GalleryBackColor = "#CCCCCC"
GalleryPhotoBorderColor = "#666666"
MenuSelectedFontColor = "#000000"
MenuSelectedColor =  "#99AABE" 
HeaderColor = "#CCCCCC"  
MenuColor = "#666666" 
MenuBorderColor = "#CCCCCC"
OutsideColor = "#F0E78C"

Dim BodyContentPaddingStyles
Select Case WEB_PAGE_ID
    Case 3, 5
        BodyContentPaddingStyles = "0 0 0 5"
    Case 4, 6
        BodyContentPaddingStyles = "0 5 0 0"
    Case Else
        BodyContentPaddingStyles = "0"
End Select
%>

<style type="text/css" media="screen">
body, a, div, td, h1, p {
	font:13px <%=FontFamily%>;
	color:<%=BodyFontColor%>;
}
body {
	scrollbar-base-color: MEDIUM;  	
	scrollbar-face-color: MEDIUM;  	
	scrollbar-track-color: BACKGROUND;  	
	scrollbar-highlight-color: LIGHT;  	
	scrollbar-3dlight-color: MEDIUM;  	
	scrollbar-shadow-color: DARK; 	
	scrollbar-darkshadow-color: MEDIUM;  	
	scrollbar-arrow-color: DARK;
}
h1 {
	text-align:left;
	font-weight:bold;
	font-size:14px;
	color:#88534F;
}
body {
	margin:3 0 0 0;
	text-align:center;
	background-color: <%=OutsideColor%>;	
}
.copyright-row td, .copyright-row a  {
	background-color:<%=OutsideColor%>;
	color:<%=LinkFontColor%>;
	font-size:10px;
	text-transform:lowercase;
}
a {  
	font-weight: normal;
	text-decoration: none;
	color:<%=LinkFontColor%>;
}
a:hover {  
	font-weight: normal;
	text-decoration: none;
	color:<%=LinkFontColor%>;
}
a:active {  
	font-weight: normal;
	text-decoration: none;
	color:<%=LinkFontColor%>;
}
a:visited {  
	font-weight: normal;
	text-decoration: none;
	color:<%=LinkFontColor%>;
}
.outer-shell, .inner-shell, .body-container {
	border-style:solid;
	border-width:1px;
	border-color:#666;
}
.outer-shell, .copyright-row {
	width:900;
}
.outer-shell {
	height:555px;
	background-color:#ccc;
}
.inner-shell {
	width:100%;
	height:100%;
	background-color:<%=MenuColor%>;
	padding:0;
}
.body-container {
	margin:5 5 3 5;
	background-color:<%=WebsiteContainerColor%>;
	padding:0;
}
.body-content {
	background-color:<%=BodyContentBackColor%>;
	width:562px;	
	padding:<%=BodyContentPaddingStyles%>;
}
.body-photo-display {
	width:264;
	height:340; /*355*/
	background-color:<%=PhotoDisplayBackColor%>;
	padding:0;
}
.body-content-liner, .body-photo-display-liner {
	width:100%;
	height:340; /*355*/
	border-style:solid;
	border-width:1px;
	border-color:#513F1B;
}
.home-photo-display {
	height:275;
	background-color:<%=PhotoDisplayBackColor%>;
	padding:0;
}
.home-photo-display-liner {
	height:275;
	border-style:solid;
	border-width:1px;
	border-color:#513F1B;
}
.body-content-liner td {
	padding:11px;
	vertical-align:top;
	height:100%;
}
.body-contact-form {
	border-style:solid;
	border-width:1px;
	border-color:#513F1B;
	vertical-align:top;
	height:100%;
	padding:10 0 10 0;
}
.body-contact-form td {
	vertical-align:top;
	height:100%;
	padding-left:10px;
}
img {
	border:0;
	padding:0;
}
input, textarea {
	border-style:solid;
	border-width:1px;
	border-color:#666;
}
</style>

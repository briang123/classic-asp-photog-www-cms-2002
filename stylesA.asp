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
Dim ImagePath

PHOTOGRAPHER_NAME = "Julie"
FontFamily = """century gothic"", sans-serif"
PhotoDisplayBackColor = "#FFFFFF"
LinkFontColor = "#FFFFFF"
BodyFontColor = "#666666"
BodyContentBackColor = "#FFFFFF"
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
OutsideColor = "#99AABE" 
%>

<style>
.Link {color:#99AABE; }
.BodyHeadline { color:<%=MenuSelectedColor%>;padding:10 0 0 0;margin:0;font:13 century gothic;font-weight:bold;}
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
	padding-left:10px;
	text-align:left;
	font-weight:normal;
	font-size:14px;
	color:<%=MenuSelectedColor%>;
}
body {
	margin:0 0 0 0;
	text-align:center;
	background-color: <%=OutsideColor%>;	
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
.thumb {
	float:left;
	height:56px;
	width:56px;
	padding:2;
	margin:5 0 0 3;
	border-style:solid;
	border-width:1px;
	border-color:#666;
	color:#666;
	background-color:#ccc;
}
.thumb a:visited, .thumb a {
	color:#000;
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
.copyright-row td {
	background-color:<%=OutsideColor%>;
	color:#eee;
	font-size:10px;
	text-transform:lowercase;
}
.inner-shell {
	width:100%;
	height:100%;
	background-color:<%=MenuColor%>;
	padding:0;
}
.header {
	height:65px;
	width:100%;
	background-color:<%=HeaderColor%>;
	border-bottom:1px solid <%=MenuColor%>;
}
.logo, .menu {
	height:100%;
}
.logo {
	text-align:left;
	width:400px;
}
.menu {
	width:492px;
}
.menu-container {
	border-style:solid none none solid;
	border-width:1px;
	border-color:<%=MenuBorderColor%>;
	width:492px;
	color:<%=MenuFontColor%>;
	background-color:<%=MenuColor%>;
}
.menu-title-over {
	border:0;
	width:492px;
	background-color:<%=MenuColor%>;
}
.menu-item, .menu-item-left, .menu-item-right {
	height:10px;
	text-transform:lowercase;
}
.menu-item-left, .menu-item-right {
	border-style:none solid none none;	
	border-color:<%=MenuBorderColor%>;	
}
.menu-item-right {
	border-color:<%=MenuColor%>;
}
.menu-item {
	border-style:none solid none none;
	border-color:<%=MenuBorderColor%>;	
}
.menu-selected {
	background-color:<%=MenuSelectedColor%>;
	color:<%=MenuSelectedFontColor%>;
}
.body-container, .body-content, .body-photo-display, .body-content-liner, .body-photo-display-liner {
	/*height:480px;*/
}
.body-container {
	margin:5 5 3 5;
	/*width:880px;*/
	
	background-color:<%=WebsiteContainerColor%>;
	padding:0;
}
.body-content {
	background-color:<%=BodyContentBackColor%>;
	width:562px;
	padding:2;
}
.body-photo-display {
	width:264px;
	height:326px;
	background-color:<%=PhotoDisplayBackColor%>;
	padding:2;
}
.body-content-liner, .body-photo-display-liner {
	width:100%;
	border-style:solid;
	border-width:1px;
	border-color:#666;
	padding:0px;
}
.body-content-liner td {
	padding:10px;
	vertical-align:top;
	height:100%;
}
img {
	border:0;
	vertical-align:top;
	padding:0;
}
input, textarea {
	border-style:solid;
	border-width:1px;
	border-color:#666;
}
</style>

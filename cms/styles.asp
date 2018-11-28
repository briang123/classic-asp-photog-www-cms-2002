<!--#include virtual="/cms/lib/__skin.asp" -->
<style>
em {
	padding:0 10 0 2;
	color:#ff0000;
}
body, td, div, span, input, textarea,th,select,p {
	font: 10 <%=CMS_NORMAL_FONT_TYPE%>; 
}
.input {
	font-size: 8pt;
}
th {
	font:10;
	font-weight:bold;
	color:#666;
	text-align:left;
	padding:0 5 2 3;
}
body {
	margin:0;
	text-align:left;
	background-color: #fff;
	color:#000;
}
.header {
	vertical-align:middle;
	height:40px;
	padding-left:30px;
	background-color:<%=abelard_top_menu_color%>;
	font-weight:bold;
	border-style:none none solid none;
	border-width:0 0 3 0;
	border-color:#fff #fff <%=abelard_border_color%> #fff;
}
.header ul {
	margin:0px;
}
li.inline {
	display:inline;
	padding:0 7 0 7;
	color:#fff;
}
li.inline-last {
	display:inline;
	padding:0 7 0 7;
	color:#fff;	
}
li.inline-last a, li.inline-last {
	font: <%=CMS_BASELINE_FONT_SIZE%> <%=CMS_MENU_LINK_FONT_TYPE%>;
	text-decoration:none;
	color:#fff;
}
li.inline a, li.inline {
	font: <%=CMS_BASELINE_FONT_SIZE%> <%=CMS_MENU_LINK_FONT_TYPE%>;
	text-decoration:none;
	color:#fff;
}
li.inline-last a {
	font:<%=CMS_BASELINE_FONT_SIZE%> <%=CMS_MENU_LINK_FONT_TYPE%>;
	font-weight:bold;
	text-decoration:underline;
	color:#fff;
}
li.inline a  {
	font: <%=CMS_BASELINE_FONT_SIZE%> <%=CMS_MENU_LINK_FONT_TYPE%>;
	font-weight:bold;	
	text-decoration:underline;
	color:#fff;
}
.title {
	background-color:<%=abelard_title_bgcolor%>;
	vertical-align:middle;
	height:20px;
	border-style:none none solid none;
	border-width:0px 0px 1px 0px;
	border-color:#ffffff #ffffff <%=abelard_border_color%> #ffffff;	
}
span {
	float:left;
}
img {
	border:0px;
}
span.page-title {
	float:none;
	font:<%=CMS_BASELINE_FONT_SIZE + 2 %> <%=CMS_TITLE_FONT_TYPE%>;
	color:#666;
}
span.sub-title {
	float:none;
	font:<%=CMS_BASELINE_FONT_SIZE + 4 %> <%=CMS_TITLE_FONT_TYPE%>;
	color:#666;
	font-weight:bold;
}
#leftnav {
	background-color:<%=abelard_left_menu_color%>;
	width:<%=CMS_LEFT_NAV_WIDTH%>;
	padding:20 2 20 0;
	vertical-align:top;
}
#leftnav a {
	padding-left:5px;
	line-height:17px;
}
#leftnav div.related-links {
	vertical-align: bottom;
	font-size:<%=CMS_BASELINE_FONT_SIZE + 2 %>;
	color:#666;
	width:100%;
	padding:5 0 3 5;
	border-style:none none solid none;
	border-width:0 0 2 0;
	border-color-bottom:#666;
}
#mainbody {
	background-color:#fff;
	width:auto;
	padding:20 10 20 10;
	vertical-align:top;
}
.menu-section-header {
	font-size:<%=CMS_BASELINE_FONT_SIZE + 6 %>;
	font-weight:bold;
	color:#666;
	border-bottom-style:solid;
	border-bottom-width:1px;
	border-bottom-color:<%=abelard_border_color%>;
}
.link-list {
	padding-top:10px;
	width:<%=CMS_HOME_PAGE_COLUMN_WIDTH%>;
}
.link-list ul, .link-list ul li{
	margin:0 0 0 10;
	padding-bottom:10px;
	list-style:none;
	list-style-image: url('<%=CMS_IMAGE_PATH %>/rect.gif');
}
p.instruction, p.admin-instruction {
	margin:5 0 0 0;
	padding:0 10 0 0;
}
p.instruction {
	width:<%=CMS_PAGE_INSTRUCTIONS_WIDTH%>;
}
table.table-container {
	padding:0;
	border:0;
	width:100%;
	height:100%;
	background-color:#fff;
}
.report td {
	padding:0 5 0 5;
	line-height:20px;
	border-bottom-style:solid;
	border-bottom-width:1px;
	border-bottom-color:#ccc;
}
.report-header {
	padding:0 5 0 5;
	font:<%=CMS_BASELINE_FONT_SIZE + 2 %> <%=CMS_TABLE_HEADER_FONT_TYPE%>;
	font-weight:bold;
	color:#666;
	background-color:#ddd;
}
.container-col1 {
	width:100px;
	text-decoration:italics;
}
.container-col2 {
	width:500px;
}
.link-list ul li a {
	text-decoration:none;
}
/*a:hover {
	text-decoration:underline;
	color:#0093B7;
}*/
.box {
	border-style:solid;
	border-color:#666;
	border-width:1px;
	padding:5px;
}
a {
	font: <%=CMS_BASELINE_FONT_SIZE + 1%> <%=CMS_MENU_LINK_FONT_TYPE%>;
	font-weight: normal;
	text-decoration: none;
}
a.menu {
	color: #000000;    
	text-transform:capitalize;
}
a.menu:visited {
	text-decoration: none;
	color: #000000;    
	text-transform:capitalize;	
}
a.menu:hover {
	text-decoration:underline;
	color: #ff0000;    
}
a.menu:active {   
	text-decoration:underline;
	color: #ff0000; 
}
.label-col {
	width:100px;
}
div.help {
	display:none;
	padding:10 10 5 10;
	border-style:dotted;
	border-color:#ccc;
	border-width:1;
	background-color:#eee;
	width:750px;	
}

</style>

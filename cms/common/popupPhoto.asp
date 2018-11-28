<html>
<head>
<title>Photo Viewing Window</title>
<style type="text/css">
<!--
body {
	margin: 0px;
	padding: 0px;
}
a img {
	border: none;
}
-->
</style>
</head>
<body onblur="self.close();">
<div><a href="javascript:self.close();"><img src="<%=Request.QueryString("Path")%>" border="0" height="<%=Request.QueryString("h")%>" width="<%=Request.QueryString("w")%>" alt="Click Image to Close"></a></div>
</body>
</html>

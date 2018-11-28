<% 
If Request.QueryString("ptype") = "GALLERY" Then
	Response.Redirect("photo-workspace.asp?gid=" & Request.QueryString("gid") & "&gln=" & Request.QueryString("gln") & "&gn=" & Request.QueryString("gn")) 
ElseIf Request.QueryString("ptype") = "CATEGORY" Then
	Response.Redirect("photo-workspace.asp?cid=" & Request.QueryString("cid") & "&cn=" & Request.QueryString("cn"))
End If

%>
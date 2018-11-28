<!-- #include virtual="/objects/cMetaData.asp" -->
<%
Dim oMeta,collMeta
Set oMeta = New cMetaData
Set collMeta = New cMetaData
collMeta.WebPageId = WEB_PAGE_ID
Call collMeta.GetMetaDataByPageId()
For Each oMeta In collMeta.MetaData.Items
	echo(vbCrLf)
	echo("<meta name=""keywords"" content=""" & QuoteCleanup(oMeta.MetaKeywords) & """>" & vbCrLf)
	echo("<meta name=""description"" content=""" & QuoteCleanup(oMeta.MetaDescription) & """>"  & vbCrLf)
	Set oMeta = Nothing
Next
Set collMeta = Nothing
%>

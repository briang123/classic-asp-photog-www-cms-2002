<%@ Language=VBScript %>
<% Option Explicit 

'--- Declarations
Dim oFileUp
Dim strFilePath
Dim strFileName
Dim oFM, oFile

'--- Instantiate the FileUp object
Set oFileUp = Server.CreateObject("SoftArtisans.FileUp")

'--- Do image path and filename lookup (from database fields or other server-side code)
'If Request.QueryString("ID") = 1 Then
'	strFilePath = Server.MapPath("sample.jpg")
'	strFileName = "sample.jpg"
'End If

If CInt(Request.QueryString("t") = 1 Then
    strFilePath = ROOT_PATH & PROOF_PATH & "/" & GetSessionVariable("PROOF_GALLERY_NAME") & "/thumbs/"
Else
    strFilePath = ROOT_PATH & PROOF_PATH & "/" & GetSessionVariable("PROOF_GALLERY_NAME") & "/"
End If
strFileName = Request.QueryString("file")

'strFilePath = Server.MapPath("C:\starktemp\secure\proofs\julie\larson\ma14.jpg")' & Request.QueryString("file"))
'strFileName = "ma14.jpg" 'Request.QueryString("file")

response.Write(strFilePath & "::" & strFileName)
response.End

'--- Use SoftArtisans.FileManager to obtain the byte-size of the file
'--- and set it in the Content-Size header
On Error Resume Next
	Set oFM = Server.CreateObject("SoftArtisans.FileManager")
	If Err.Number <> 0 Then
		Response.Write "<B>Error creating FileManager object.</B> FileManager must " & _
						"be installed for this sample to run."
		Response.End
	End If
On Error Goto 0

On Error Resume Next
	Set oFile = oFM.GetFile(strFilePath)
	If Err.Number <> 0 Then
		Response.Write "<B>FileManager could not open the file at:</B> " & strFilePath & _
						"<BR>" & Err.Description & " (" & Err.Source & ")"
		Response.End
	End If
On Error Goto 0	
	
'--- Set response headers
'--- Set the ContentType as appropriate
Response.ContentType = "image/jpeg"

'--- Set Content-Disposition to "inline" and specify the filename.
Response.AddHeader "Content-Disposition", "inline;filename=""" & strFileName & """"
Response.AddHeader "Content-Length", oFile.Size
	
'--- Send the file
On Error Resume Next
	oFileUp.TransferFile strFilePath
	If Err.Number <> 0 Then
		Response.Clear
		Response.Write "<B>FileUp could not download the file at: </B> " & strFilePath & _
						"<BR>" & Err.Description & " (" & Err.Source & ")"
		Response.End
	End If
On Error Goto 0
	
	
'--- Clean up
Set oFile = Nothing
Set oFM = Nothing
Set oFileUp = Nothing
Response.End
%>


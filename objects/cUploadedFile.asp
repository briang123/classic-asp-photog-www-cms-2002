<%
Class cUploadedFile
    Public ContentType
    Public FileName
    Public FileData
    
    Public Property Get FileSize()
        FileSize = LenB(FileData)
    End Property

    Public Sub SaveToDisk(sPath)
        Dim oFS, oFile
        Dim nIndex
    
        If sPath = "" Or FileName = "" Then Exit Sub
        If Mid(sPath, Len(sPath)) <> "\" Then sPath = sPath & "\"
    
        Set oFS = Server.CreateObject("Scripting.FileSystemObject")
        If Not oFS.FolderExists(sPath) Then Exit Sub
        
        Set oFile = oFS.CreateTextFile(sPath & FileName, True)
        
        For nIndex = 1 to LenB(FileData)
            oFile.Write Chr(AscB(MidB(FileData,nIndex,1)))
        Next

        oFile.Close
    End Sub
    
    Public Sub SaveToDatabase(ByRef oField)
        If LenB(FileData) = 0 Then Exit Sub
        
        If IsObject(oField) Then
            oField.AppendChunk FileData
        End If
    End Sub
End Class
%>
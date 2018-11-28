
<%
'############ Common Functions, used unchanged by many classes
'Takes an SQL String
'Calls other functions

'Takes an SQL Query
'Runs the Query and returns a recordset
Function LoadRSFromDB(p_strSQL)
'on error resume next
    dim rs, cmd

    Set rs = Server.CreateObject("adodb.Recordset")
    Set cmd = Server.CreateObject("adodb.Command")
    
    'Run the SQL
    cmd.ActiveConnection  = CONNECTION_STRING
    cmd.CommandText = p_strSQL '"SELECT c.ConfigId, c.ConfigKey, c.ConfigValue, c.WebsiteId, c.ActiveFlag FROM tblConfig c, tblWebsite w WHERE c.WebsiteId = w.SiteId ORDER BY c.WebsiteId, c.ConfigKey ASC" 
    cmd.CommandType = adCmdText
    cmd.Prepared = true

'		response.write p_strSQL & "<br><br>"
		
    rs.CursorLocation = adUseClient
    rs.Open cmd, , adOpenForwardOnly, adLockReadOnly

    if Err <> 0 then
        Err.Raise  Err.Number, "ADOHelper: RunSQLReturnRS", Err.Description
    end if

    ' Disconnect the recordsets and cleanup  
    Set rs.ActiveConnection = Nothing  
    Set cmd.ActiveConnection = Nothing
    Set cmd = Nothing
    Set LoadRSFromDB = rs	
End Function

Function RunSQL(ByVal p_strSQL)
	' Create the ADO objects
	Dim cmd
	Set cmd = Server.CreateObject("adodb.Command")

	cmd.ActiveConnection  = CONNECTION_STRING
	cmd.ActiveConnection.BeginTrans
	cmd.CommandText = p_strSQL
	cmd.CommandType = adCmdText
    
    'response.Write(p_strSQL)
    'response.End
	' Execute the query without returning a recordset
	' Specifying adExecuteNoRecords reduces overhead and improves performance
	cmd.Execute true, , adExecuteNoRecords
	cmd.ActiveConnection.CommitTrans

	if Err <> 0 then
		cmd.ActiveConnection.RollBackTrans
		Err.Raise  Err.Number, "ADOHelper: RunSQL", Err.Description
	end if

	' Cleanup
	Set cmd.ActiveConnection = Nothing
	Set cmd = Nothing
End Function

Function InsertRecord(tblName, strAutoFieldName, ArrFlds, ArrValues )
	dim conn, rs, thisID   
    Set conn = Server.CreateObject ("ADODB.Connection")
    Set rs = Server.CreateObject ("ADODB.Recordset")
	
    conn.open CONNECTION_STRING
    conn.BeginTrans
    rs.Open tblName, conn, adOpenKeyset, adLockOptimistic, adCmdTable

    rs.AddNew  ArrFlds, ArrValues
    rs.Update 

    thisID = rs(strAutoFieldName)

    rs.Close
    Set rs = Nothing

    conn.CommitTrans        
    conn.close
    Set conn = Nothing

    If Err.number = 0 Then
        InsertRecord = thisID
    End If         
End Function

function SingleQuotes(pStringIn)
    if pStringIn = "" or isnull(pStringIn) then 
	SingleQuotes = NULL
	exit function
    end if
	Dim pStringModified
    pStringModified = Replace(pStringIn,"'","''")
    SingleQuotes =  pStringModified
end function

public function echo(p_STR)
    response.write p_Str
end function

public function die(p_STR)
    echo p_Str
    response.end
end function

public function echobr(p_STR)
    echo p_Str & "<br>" & vbCRLF
end function

public function htmlencode(p_STR)
    htmlencode = trim(server.htmlencode(p_Str & " "))
end function

Randomize 'Insure that the numbers are really random
    Function RandomString(p_NumChars)
    Dim n
    Dim tmpChar,tmpString
    for n = 0 to p_NumChars
        tmpChar = Chr(Int(32+( Rnd * (126-33))))
        'Random characters (letters, numbers, etc.)
        tmpString = tmpString & tmpChar
    next
    RandomString = tmpString
End Function


%>
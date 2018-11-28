<% 
'----------------------------------------------------------------------
' FUNCTION NAME:	CloseCmd
' CREATED BY:		Brian Gaines
' CREATED ON:		11/5/2002
' INPUTS:			cmd - command object
' PURPOSE:			Closes the command object
'----------------------------------------------------------------------
function CloseCmd(cmd)
	if Not cmd.State = adStateOpen Then
		If cmd.State = adStateOpen Then
			Set cmd.ActiveConnection = Nothing
		End If
		Set cmd = Nothing
	End If
end function

'----------------------------------------------------------------------
' FUNCTION NAME:	CloseDBConn
' CREATED BY:		Brian Gaines
' CREATED ON:		11/5/2002
' INPUTS:			conn - connection object
' PURPOSE:			Closes the database connection object, which will
'					free up the connection object to the pool
'----------------------------------------------------------------------
function CloseDBConn(conn)
	If conn.State = adStateOpen Then
		conn.Close
	End If
end function

'----------------------------------------------------------------------
' FUNCTION NAME:	CloseRS
' CREATED BY:		Brian Gaines
' CREATED ON:		11/5/2002
' INPUTS:			recSet
' PURPOSE:			Closes recordset object if open
'----------------------------------------------------------------------
function CloseRS(recSet)
	If Not recSet Is Nothing Then
		If recSet.State = adStateOpen Then
			  recSet.Close
		End	If
		Set recSet = Nothing
	End If
end function

%>
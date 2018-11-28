<%
Dim rs
Dim dbConn

Function GetConnection()

	On Error Resume Next
	
	'----------------------------------------------
	' Check if reference to database connection 
	' exists and assign a reference to it.
	'----------------------------------------------	

	If Not IsObject(dbConn) Then
		Call OpenDB()
	End If

	GetConnection = dbConn		
	
	HandleErrors "DATABASE CONNECTION"
	
End Function

Function OpenDB()

	'On Error Resume Next

	'----------------------------------------------
	' Open the database connection object
	'----------------------------------------------	

	If Not IsObject(dbConn) Then
		Set dbConn = Server.CreateObject("ADODB.Connection")
		
		dbConn.Open CONNECTION_STRING
		
		HandleErrors "DATABASE CONNECTION"
	End If

End Function

Function CloseDB()

	On Error Resume Next

	'----------------------------------------------
	' Close the database connection object
	'----------------------------------------------	
	
	If UCase(TypeName(dbConn)) = "CONNECTION" then
		dbConn.Close
		
		Set dbConn = Nothing
		
		HandleErrors "DATABASE CONNECTION"
	End If
	
End Function

Function OpenRS(objRS)

	'----------------------------------------------
	' Open the recordset ojbect
	'----------------------------------------------	

	OpenDB()
	
	Set objRS = Server.CreateObject("ADODB.Recordset")
	
	HandleErrors "RECORDSET OBJECT"	
	
End Function

Function CloseRS(objRS)
	
	On Error Resume Next

	'----------------------------------------------
	' Close the recordset object
	'----------------------------------------------	

	Set objRS = Nothing
	HandleErrors "RECORDSET OBJECT"

End Function

Sub HandleErrors(ByVal strErrType)
	On Error Resume Next
	If Err.number <> 0 then			
'		Session(APPVARNAME & "ErrorNumber") = Err.number
'		Session(APPVARNAME & "ErrorDescription") = Err.description
'		Session(APPVARNAME & "ErrorSource") = Err.Source
'		Session(APPVARNAME & "ErrorType") = strErrType

		response.write err.number & err.description
'		Response.Redirect(ERRORPAGEPATH)
	End if

End Sub

Function CreateDictionary(objDict)

	On Error Resume Next
	
	Set objDict = Server.CreateObject("Scripting.Dictionary")

	'----------------------------------------------
	' Open the Sql statement storage object
	'----------------------------------------------	

	If IsObject(objDict) Then
		CreateDictionary = True
	Else
		CreateDictionary = False
	End If

	HandleErrors "OPEN DICTIONARY OBJECT"

End Function

Function CloseDictionary(objDict)

	On Error Resume Next

	'----------------------------------------------
	' Close the Sql statement storage object
	'----------------------------------------------	
	
	If IsObject(objDict) Then
		Set objDict = Nothing
		CloseDictionary = True
	Else
		CloseDictionary = False
	End If

	HandleErrors "CLOSE DICTIONARY OBJECT"
		
End Function

Function RemoveSqlFromQueue(objDict,key)

	On Error Resume Next

	'----------------------------------------------
	' Remove Sql statements from the queue
	'----------------------------------------------	
	
	If IsObject(objDict) Then
	
		'----------------------------------------------
		' Remove all Sql statements from queue
		'----------------------------------------------	

		If StringEmptyOrNull(key) Then
			objDict.RemoveAll
			If CloseDictionary(objDict) Then
				RemoveSqlFromQueue = True
			Else
				RemoveSqlFromQueue = False	
			End If
		Else
		
			'----------------------------------------------
			' Remove specific Sql statement from queue
			'----------------------------------------------	
				
			objDict.Remove(key)
			RemoveSqlFromQueue = True
		End If
	Else
		RemoveSqlFromQueue = False
	End If

	HandleErrors "REMOVE SQL FROM DICTIONARY OBJECT"
End Function

Function AppendToSqlQueue(objDict,strSql)

	On Error Resume Next

	'----------------------------------------------
	' Add Sql statements to the queue
	'----------------------------------------------	
	
	If IsObject(objDict) Then
		objDict.Add objDict.Count, strSql	
		HandleErrors "APPEND SQL TO DICTIONARY OBJECT"
		AppendToSqlQueue = True
	Else
		AppendToSqlQueue = False
	End If

End Function

Function ExecuteSqlQueue(objDict)
	
	On Error Resume Next

	Call OpenDb()

	Dim collItems, intLoop, sql
	
	'----------------------------------------------
	' Execute a consecutive list of Sql statements
	'----------------------------------------------		
	
	ExecuteSqlQueue = False	

	'----------------------------------------------
	' If the queue and connection object exist
	'----------------------------------------------		

	If IsObject(objDict) Then
	
		If IsObject(dbConn) Then
	
			'----------------------------------------------------------------	
			' Start the ADO transaction object then execute Sql statements
			'----------------------------------------------------------------			
	
			dbConn.BeginTrans
	
			collItems = objDict.Items
			
			'----------------------------------------------			
			' Loop through queue and execute Sql statements
			'----------------------------------------------			
			
			For intLoop = 0 to objDict.Count - 1
'			response.write collItems(intLoop) & "<BR>"
			
				dbConn.Execute collItems(intLoop),,adCmdText + adExecuteNoRecords

				If Err.Number <> 0 Then
					dbConn.RollbackTrans
					HandleErrors "EXECUTE SQL IN DICTIONARY OBJECT"
				End If
				
			Next

			'----------------------------------------------
			' Commit ADO transaction if successful
			'----------------------------------------------				

			dbConn.CommitTrans

			ExecuteSqlQueue = True

		End If

	End If
	
End Function

Function ExecuteSql(sql)

	On Error Resume Next
	
	Call OpenDb()

	'----------------------------------------------
	' Execute Sql statement in an ADO transaction
	'----------------------------------------------
		
	dbConn.BeginTrans
	
	ExecuteSql = False	

	'----------------------------------------------
	' Check if we created connection object
	' Execute Sql
	' Rollback changes if there is an error
	'----------------------------------------------

	If IsObject(dbConn) Then
		dbConn.Execute sql,,adCmdText + adExecuteNoRecords
		If Err.Number <> 0 Then
			dbConn.RollbackTrans
			HandleErrors "EXECUTE SQL STATEMENT BATCH"
		End If
		ExecuteSql = True
	End If

	'----------------------------------------------
	' Commit transaction if successful executing
	'----------------------------------------------	

	dbConn.CommitTrans
	
End Function

Function DeleteRecord(blnDeleteSingleRecord,strTableName,strPrimaryKey,idToDelete)

	Dim deleteSql
	DeleteRecord = False

	'-----------------------------------
	' Verify arguments
	'-----------------------------------
	If StringNotEmptyOrNull(blnDeleteSingleRecord) And StringNotEmptyOrNull(strTableName) And _
		StringNotEmptyOrNull(strPrimaryKey) And StringNotEmptyOrNull(idToDelete) Then


		'make a call to the BuildDeleteStatement function to get this
	
		deleteSql = "DELETE FROM " & strTableName & " WHERE " & strPrimaryKey
	
		If blnDeleteSingleRecord Then

			'-----------------------------------
			' Deleting only 1 record
			'-----------------------------------
		
			deleteSql = deleteSql & " = " & CInt(idToDelete)
	
		Else
		
			'-----------------------------------
			' Deleting list of records
			'-----------------------------------
		
			deleteSql = deleteSql & " IN (" & idToDelete & ")"
			
		End If
	
		'---------------------------------------------------
		' Execute the delete statement and return message
		'---------------------------------------------------
	
		If ExecuteSql(deleteSql) Then	
			strMessage = FormatMessage(MSG_SUCCESSFUL_DELETE)
			DeleteRecord = True		
		Else
			strMessage = FormatMessage(MSG_DELETE_FAILURE)
		End If

	End If

'	Else	

		'-----------------------------------
		' Deleting batch of records
		'-----------------------------------
	
'		Dim counter, objDictionary			' 

'		Call CreateDictionary(objDictionary)

		'-------------------------------------------------------------
		' Loop through each checkbox value and add sql to queue
		'-------------------------------------------------------------
		
'		For counter = LBound(idList) To UBound(idList)
'			deleteSql = "DELETE FROM " & strTableName & " WHERE " & strPrimaryKey & " IN (" & idToDelete & ")"
'			Call AppendToSqlQueue(objDictionary,deleteSql)		
'		Next 

		'---------------------------------------------
		' Execute all sql in queue then do cleanup
		'---------------------------------------------
		
'		If ExecuteSqlQueue(objDictionary) Then
		
'			If RemoveSqlFromQueue(objDictionary,"") Then
'				strMessage = FormatMessage(MSG_SUCCESSFUL_DELETE)
'				DeleteRecord = True
'			End If
			
'		End If
		
'		Call CloseDictionary(objDictionary)

'	End If

End Function

Function BuildDeleteStatement(strTableName,strPrimaryKey,idToDelete)

	Dim deleteSql
	BuildDeleteStatement = ""

	'-----------------------------------
	' Verify arguments
	'-----------------------------------
	If StringNotEmptyOrNull(strTableName) And StringNotEmptyOrNull(strPrimaryKey) And _
		StringNotEmptyOrNull(idToDelete) Then

		deleteSql = "DELETE FROM " & strTableName & " WHERE " & strPrimaryKey
	
		If Instr(idToDelete,",") > 0 Then

			'-----------------------------------
			' Deleting list of records
			'-----------------------------------
		
			deleteSql = deleteSql & " IN (" & idToDelete & ")"
	
		Else

			'-----------------------------------
			' Deleting only 1 record
			'-----------------------------------
		
			deleteSql = deleteSql & " = " & CInt(idToDelete)
					
		End If
	
		BuildDeleteStatement = deleteSql

	End If

End Function


Function RecordExists(checkSql)

	'----------------------------------------------------------------
	' Get existance of a record in a particular table
	'----------------------------------------------------------------

	Call OpenRs(rs)
	rs.Open checkSql, dbConn

	rs.MoveFirst
	' determine how many rows were returned	
	If rs(0) > 0 Then
		RecordExists = True
	Else
		RecordExists = False
	End If
	
	Call CloseRs(rs)

End Function

Function GetSingleValue(tbl,fld,keyname,keyvalue)
	
	Dim sql
	Dim rsLocal
	
	'----------------------------------------------------------------
	' Get single value from any given table specified as argument
	'----------------------------------------------------------------
	
	sql = "SELECT " & fld & " FROM " & tbl & " WHERE " & keyname & " = " & keyvalue
	
	Call OpenRs(rsLocal)
	
	rsLocal.Open sql, dbConn	
	If Not rsLocal.BOF And Not rsLocal.EOF Then
		rsLocal.MoveFirst	
		GetSingleValue = rsLocal(0)
	Else
		GetSingleValue = ""
	End If
	Call CloseRs(rsLocal)

End Function

Function GetRecordsetMoveFirst(objRs,lclSql)

	objRs.Open lclSql, dbConn, adOpenDynamic
	
	If objRs.BOF And objRs.EOF Then
		strMessage = FormatMessage(MSG_NO_RECORDS_FOUND)
	Else
		objRs.MoveFirst
	End If
	
End Function

Function GetRowCount(tbl,whereClause)
	
	Dim sql
	Dim rsLocal
		
	sql = "SELECT COUNT(*) FROM " & tbl & " WHERE " & whereClause
	
	Call OpenRs(rsLocal)
	
	rsLocal.Open sql, dbConn	
	If Not rsLocal.BOF And Not rsLocal.EOF Then
		rsLocal.MoveFirst	
		GetRowCount = rsLocal(0)
	Else
		GetRowCount = ""
	End If
	Call CloseRs(rsLocal)

End Function

%>
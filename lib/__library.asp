<%
'--------------------------------------------------------------------------------------
' CREATED BY: BRIAN GAINES
' FILENAME: 	__library.asp
' PURPOSE:		Stored common routines for the Gaines Consulting web application		
'--------------------------------------------------------------------------------------

' PURPOSE: 		Builds the combo box with list of values from database and selects the appropriate client
'							name based on the DeltaValue passed into this function.			
' INPUTS: 		DeltaValue - option value to select in the combo box
Function ClientList(DeltaValue)
	
	dim combo

	' open recordset to get the clientId and client name
	Call openRS(rs)
	rs.Open "SELECT ClientID, ClientName FROM tblClient ORDER BY ClientName ASC", GetConnection
	
	' create the combo box by passing recordset object, name of combo box, option value, option text, and the value to match up with
	combo = BuildComboBox(rs,"selClientList",DeltaValue)

	' free up the recordset
	Call CloseRS(rs)

	ClientList = combo

End Function


' PURPOSE: 		Builds the combo box with list of values from database and selects the appropriate company
'				name based on the DeltaValue passed into this function.
' INPUTS: 		DeltaValue - option value to select in the combo box	
Function CompanyList(DeltaValue)

	dim combo

	' open recordset to get the clientId and client name
	Call openRS(rs)
	rs.Open "SELECT CompanyID, CompanyName FROM tblCompanies ORDER BY CompanyName ASC", GetConnection
	
	' create the combo box by passing recordset object, name of combo box, option value, option text, and the value to match up with
	combo = BuildComboBox(rs,"selHiringFirm",DeltaValue)

	' free up the recordset
	Call CloseRS(rs)

	CompanyList = combo
	
End Function

' PURPOSE:		Builds a order list that is numbered from 1 to the (number of projects + 1) created in
'							the projects table, which allows the user to select in what order to display the project info.
' INPUTS: 		DeltaValue - option value to select in the combo box	
Function OrderList(DeltaValue)

	Dim strList, inCurrOrderNum, intNextOrderNum

	' build the combo box list. Calls a function GetMaxId to get the latest id for a particular table.
	strList = "<SELECT NAME=" & DblQt("selOrder") & ">" & vbCRLF
	intNextOrderNum = GetMaxId("tblProjects","projectID") + 1
	for intCurrOrderNum = 1 to intNextOrderNum
		strList = strList & "<OPTION VALUE=" & DblQt(intCurrOrderNum) & setComboOption(intCurrOrderNum,DeltaValue,1) & ">" & intCurrOrderNum & "</OPTION>" & vbCRLF
	next
	strList = strList & "</SELECT>" & vbCRLF

	' return the combo box back to caller
	OrderList = strList

End Function

' PURPOSE: 		Builds an Active Flag list and selects the appropriate value based on the DeltaValue
' INPUTS: 		DeltaValue - option value to select in the combo box
Function ActiveList(DeltaValue)

	Dim strList	

	' build the combo box list.
	strList = "<SELECT NAME=" & DblQt("selActive") & ">" & vbCRLF
	strList = strList & "<OPTION VALUE=" & DblQt(-1) & setComboOption(-1,Cint(DeltaValue),1) & "> -- SELECT ONE -- </OPTION>" & vbCRLF
	strList = strList & "<OPTION VALUE=" & DblQt(0) & setComboOption(0,Cint(DeltaValue),1) & ">InActive</OPTION>" & vbCRLF
	strList = strList & "<OPTION VALUE=" & DblQt(1) & setComboOption(1,Cint(DeltaValue),1) & ">Active</OPTION>" & vbCRLF		
	strList = strList & "</SELECT>" & vbCRLF

	' return the combo box back to caller
	ActiveList = strList

End Function

Function StatusList(DeltaValue)

	dim combo

	' open recordset to get the clientId and client name
	Call openRS(rs)
	rs.Open "SELECT statusId, status FROM tblStatus ORDER BY statusorder ASC", GetConnection
	
	' create the combo box by passing recordset object, name of combo box, option value, option text, and the value to match up with
	combo = BuildComboBox(rs,"selStatus",DeltaValue)

	' free up the recordset
	Call CloseRS(rs)

	StatusList = combo
	
End Function

Function UserList(DeltaValue,comboName,Company)

	dim combo

	' open recordset to get the clientId and client name
	Call openRS(rs)
	rs.Open 	"select id, firstname, lastname from tblUsers u where company = '" & Company & "' order by lastname asc", GetConnection
	
	' create the combo box by passing recordset object, name of combo box, option value, option text, and the value to match up with
	combo = BuildComboBox(rs,comboName,DeltaValue)

	' free up the recordset
	Call CloseRS(rs)

	UserList = combo
	
End Function

Function ProjectList(DeltaValue,Company,Role)

	dim combo

	' open recordset to get the clientId and client name
	Call openRS(rs)
	
'	If Role = "Master" Then
'		rs.Open 	"select projId, project from tblProjects order by projId asc", GetConnection	
'	Else
		rs.Open 	"select projId, project from tblProjects where company = " & makeString(Company) & " order by projId asc", GetConnection
'	End If
	
	' create the combo box by passing recordset object, name of combo box, option value, option text, and the value to match up with
	combo = BuildComboBox(rs,"selProject",DeltaValue)

	' free up the recordset
	Call CloseRS(rs)

	ProjectList = combo
	
End Function

Function ProjectCompanyList(DeltaValue)

	dim combo

	' open recordset to get the clientId and client name
	Call openRS(rs)
	rs.Open 	"select distinct(company) from tblUsers order by company asc", GetConnection	
	
	' create the combo box by passing recordset object, name of combo box, option value, option text, and the value to match up with
	Dim strList	
	strList = "<SELECT NAME=" & DblQt("selCompany") & ">" & vbCRLF
	With rs
		.MoveFirst
		do while not .EOF
			strList = strList & "<OPTION VALUE=" & DblQt(rs(0)) & setComboOption(rs(0),DeltaValue,1) & ">" & rs(0) & "</OPTION>" & vbCRLF
			.MoveNext
		loop
	End With
	strList = strList & "</SELECT>" & vbCRLF

	' free up the recordset
	Call CloseRS(rs)

	ProjectCompanyList = strList
	
End Function

Function RoleList(DeltaValue)
	
	dim combo

	Call openRS(rs)
	rs.Open "SELECT RoleId, Role FROM tblRole ORDER BY RoleId ASC", GetConnection
	
	' create the combo box by passing recordset object, name of combo box, option value, option text, and the value to match up with
	combo = BuildComboBox(rs,"selRole",DeltaValue)

	' free up the recordset
	Call CloseRS(rs)

	RoleList = combo

End Function

Function BatchDateList(DeltaValue,Proj)
	
	dim combo

	Call openRS(rs)
	rs.Open "SELECT DISTINCT(Batch) FROM tblRequests WHERE status <> 6 AND batch <> #12/31/1899# AND projId = " & Proj & " ORDER BY Batch ASC", GetConnection
	
	' create the combo box by passing recordset object, name of combo box, option value, option text, and the value to match up with
	'combo = BuildComboBox(rs,"selBatch",DeltaValue)

	Dim strList	

	strList = "<SELECT NAME=" & DblQt("selBatch") & ">" & vbCRLF
	if DeltaValue = "12/31/1899" then
		strList = strList & "<OPTION VALUE=" & DblQt("12/31/1899") & " SELECTED>NO BATCH ASSIGNED</OPTION>" & vbCRLF
	end if
	With rs
		.MoveFirst
		do while not .EOF
			strList = strList & "<OPTION VALUE=" & DblQt(rs(0)) & setComboOption(rs(0),DeltaValue,1) & ">" & rs(0) & "</OPTION>" & vbCRLF
			.MoveNext
		loop
	End With
	strList = strList & "</SELECT>" & vbCRLF

	' free up the recordset
	Call CloseRS(rs)

	BatchDateList = strList

End Function

Function BuildRequestActionQueryString(sid,uid,pg,id)
	Dim qstr,bAmp
	bAmp = False
	qstr = ""
	
	If Not IsEmpty(sid) And IsNumeric(sid) Then
		qstr = qstr & "?status=" & sid
		bAmp = True
	End If
	
	If Not IsEmpty(uid) And IsNumeric(uid) Then
		if bAmp = True then
			qstr = qstr & "&user=" & uid
		else
			qstr = qstr & "?user=" & uid
			bAmp = True
		end if		
	End If

	If Not IsEmpty(pg) And IsNumeric(pg) Then
		if bAmp = True then
			qstr = qstr & "&page=" & pg
		else
			qstr = qstr & "?page=" & pg
			bAmp = True
		end if		
	End If

	If Not IsEmpty(id) And IsNumeric(id) Then
		if bAmp = True then
			qstr = qstr & "&id=" & id
		else
			qstr = qstr & "?id=" & id		
			bAmp = True
		end if
	End If

	If Not IsEmpty(qstr) Then
		BuildRequestActionQueryString = qstr
	End If
	
End Function
%>
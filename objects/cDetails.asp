<%
Class cDetails

	Private m_ID
	Private m_Details
	Private m_DetailsText
	Private m_DetailsCount
	
	Sub Class_Initialize()
		Set m_Details = Server.CreateObject("Scripting.Dictionary")
		Me.ID = 0
	End Sub
	
	Sub Class_Terminate()
		Set m_Details = Nothing
	End Sub
	
	Public Property Get IsError
		IsError = eval(Me.ID < 0)
	End Property
		
	Public Property Get Details()
		Set Details = m_Details
	End Property
		
	Public Property Let ID(p_Value)
		m_ID = p_Value
	End Property
	Public Property Get ID()
		ID = CLng(m_ID)
	End Property	

	Public Property Let DetailsText(p_Value)
		m_DetailsText = p_Value
	End Property
	Public Property Get DetailsText()
		DetailsText = m_DetailsText
	End Property
	
	Public Property Let DetailsCount(p_Value)
		m_DetailsCount = p_Value
	End Property
	Public Property Get DetailsCount()
		DetailsCount = CLng(m_DetailsCount)
	End Property
	
	Public Function AddDetails()
        If Me.ID < 1 then
            Dim arr1, arr2
            arr1 = Array("DetailsText")
            arr2 = Array(SingleQuotes(Me.DetailsText))
            Me.ID = InsertRecord("tblDetails", "DetailsId", arr1, arr2)            
        End if
		AddDetails = Eval(Not Me.IsError)
	End Function
	
	Public Function UpdateDetails()
	    Dim strSQL
		strSQL = strSQL & " UPDATE tblDetails SET "
		strSQL = strSQL & " DetailsText = '" & SingleQuotes(Me.DetailsText) & "'"
		strSQL = strSQL & " WHERE DetailsId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function DeleteDetails()
        Dim strSQL
		strSQL = "DELETE * FROM tblDetails WHERE DetailsId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function GetDetailsById()
        Dim strSQL
        strSQL = "SELECT DetailsId, DetailsText FROM tblDetails WHERE DetailsId= " & Me.ID    
        FillObjectFromRS(strSQL)
	End Function
	
	Public Function GetDetails()
        Dim strSQL
        strSQL = "SELECT DetailsId, DetailsText FROM tblDetails"
        FillObjectFromRS(strSQL)
	End Function
	
	'***********************************************
	' PRIVATE METHODS
	'***********************************************
    Private Function FillObjectFromRS(p_strSQL)

        Dim rs
        Set rs = LoadRSFromDB(p_strSQL)		
		
        Dim oDetails
		Dim counter
		counter = 0
		Do While Not rs.EOF
			counter = counter + 1
			Set oDetails = New cDetails		
			oDetails.ID = rs("DetailsId")
			oDetails.DetailsText = QuoteCleanup(rs("DetailsText"))
			m_Details.Add oDetails.ID, oDetails
            rs.MoveNext
        Loop
		Me.DetailsCount = counter
		rs.Close
        Set rs = Nothing
    End Function

    Private Function LoadData(p_strSQL)
        Dim rs
        Set rs = LoadRSFromDB(p_strSQL)
        LoadData = FillFromRS(rs)
        rs.Close
        Set rs = Nothing
    End Function

End Class
%>


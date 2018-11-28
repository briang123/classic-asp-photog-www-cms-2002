<%
Class cAbout

	Private m_ID
	Private m_Abouts
	Private m_AboutText
	Private m_AboutCount
	
	Sub Class_Initialize()
		Set m_Abouts = Server.CreateObject("Scripting.Dictionary")
		Me.ID = 0
	End Sub
	
	Sub Class_Terminate()
		Set m_Abouts = Nothing
	End Sub
	
	Public Property Get IsError
		IsError = eval(Me.ID < 0)
	End Property
		
	Public Property Get Abouts()
		Set Abouts = m_Abouts
	End Property
		
	Public Property Let ID(p_Value)
		m_ID = p_Value
	End Property
	Public Property Get ID()
		ID = CLng(m_ID)
	End Property	

	Public Property Let AboutText(p_Value)
		m_AboutText = p_Value
	End Property
	Public Property Get AboutText()
		AboutText = m_AboutText
	End Property
	
	Public Property Let AboutCount(p_Value)
		m_AboutCount = p_Value
	End Property
	Public Property Get AboutCount()
		AboutCount = CLng(m_AboutCount)
	End Property
	
	Public Function AddAboutText()
        If Me.ID < 1 then
            Dim arr1, arr2
            arr1 = Array("AboutText")
            arr2 = Array(Me.AboutText)
            Me.ID = InsertRecord("tblAbout", "AboutId", arr1, arr2)            
        End if
		AddAboutText = Eval(Not Me.IsError)
	End Function
	
	Public Function UpdateAboutText()
	    Dim strSQL
		strSQL = strSQL & " UPDATE tblAbout SET "
		strSQL = strSQL & " AboutText = '" & SingleQuotes(Me.AboutText) & "'"
		strSQL = strSQL & " WHERE AboutId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function DeleteAboutText()
        Dim strSQL
		strSQL = "DELETE * FROM tblAbout WHERE AboutId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function GetAboutTextById()
        Dim strSQL
        strSQL = "SELECT AboutId, AboutText FROM tblAbout WHERE AboutId= " & Me.ID    
        FillObjectFromRS(strSQL)
	End Function
	
	Public Function GetAboutText()
        Dim strSQL
        strSQL = "SELECT AboutId, AboutText FROM tblAbout"
        FillObjectFromRS(strSQL)
	End Function
	
	'***********************************************
	' PRIVATE METHODS
	'***********************************************
    Private Function FillObjectFromRS(p_strSQL)

        Dim rs
        Set rs = LoadRSFromDB(p_strSQL)		
		
        Dim oAbout
		Dim counter
		counter = 0
		Do While Not rs.EOF
			counter = counter + 1
			Set oAbout = New cAbout		
			oAbout.ID = rs("AboutId")
			oAbout.AboutText = QuoteCleanup(rs("AboutText"))
			m_Abouts.Add oAbout.ID, oAbout
            rs.MoveNext
        Loop
		Me.AboutCount = counter
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


<%
Class cMetaData

	Private m_ReturnCode
	Private m_ID
	Private m_WebPageId
	Private m_WebPage
	Private m_MetaKeywords
	Private m_MetaDescription
	Private m_MetaData
	Private m_MetaCount
	
	Sub Class_Initialize()
		Set m_MetaData = Server.CreateObject("Scripting.Dictionary")
		Me.ID = 0
	End Sub
	
	Sub Class_Terminate()
		Set m_MetaData = Nothing
	End Sub
	
	Public Property Get IsError
		IsError = eval(Me.ID < 0)
	End Property
		
	Public Property Get MetaData()
		Set MetaData = m_MetaData
	End Property
		
	Public Property Get ReturnCode()
		ReturnCode = m_ReturnCode
	End Property
	
	Public Property Let ID(p_Value)
		m_ID = p_Value
	End Property
	Public Property Get ID()
		ID = m_ID
	End Property
	
	Public Property Let WebPageId(p_Value)
		m_WebPageId = p_Value
	End Property
	Public Property Get WebPageId
		WebPageId = m_WebPageId
	End Property
	
	Public Property Let WebPage(p_Value)
		m_WebPage = p_Value
	End Property
	Public Property Get WebPage
		WebPage = m_WebPage
	End Property

	Public Property Let MetaKeywords(p_Value)
		m_MetaKeywords = p_Value
	End Property
	Public Property Get MetaKeywords
		MetaKeywords = m_MetaKeywords
	End Property

	Public Property Let MetaDescription(p_Value)
		m_MetaDescription = p_Value
	End Property
	Public Property Get MetaDescription
		MetaDescription = m_MetaDescription
	End Property

	Public Property Let MetaCount(p_Value)
		m_MetaCount = p_Value
	End Property
	Public Property Get MetaCount()
		MetaCount = m_MetaCount
	End Property
	
	Public Function AddMetaData()
		If Me.MetaKeywords = "" Then 
			Me.MetaKeywords = Null
		Else
			Me.MetaKeywords = SingleQuotes(Me.MetaKeywords)
		End If
		
		If Me.MetaDescription = "" Then
			Me.MetaDescription = Null
		Else
			Me.MetaDescription = SingleQuotes(Me.MetaDescription)
		End If
		
        If Me.ID < 1 then
            Dim arr1, arr2
            arr1 = Array("WebPage", "MetaKeywords", "MetaDescription")
            arr2 = Array(Me.WebPageId, Me.MetaKeywords, Me.MetaDescription)
            Me.ID = InsertRecord("tblMetaData", "MetaId", arr1, arr2)            
        End if
		AddMetaData = Eval(Not Me.IsError)
	End Function
	
	Public Function UpdateMetaData()
		If Me.MetaKeywords = "" Then 
			Me.MetaKeywords = Null
		Else
			Me.MetaKeywords = SingleQuotes(Me.MetaKeywords)
		End If
		
		If Me.MetaDescription = "" Then
			Me.MetaDescription = Null
		Else
			Me.MetaDescription = SingleQuotes(Me.MetaDescription)
		End If
		
	    Dim strSQL
		strSQL = strSQL & " UPDATE tblMetaData SET "
		strSQL = strSQL & " WebPage = " & Me.WebPageId & ","
		strSQL = strSQL & " MetaKeywords = '" & SingleQuotes(Me.MetaKeywords) & "', "
		strSQL = strSQL & " MetaDescription = '" & SingleQuotes(Me.MetaDescription) & "' " 
		strSQL = strSQL & " WHERE MetaId = " & Me.ID
		RunSQL strSQL            
		
		UpdateMetaData = Eval(Not Me.IsError)
	End Function
	
	Public Function DeleteMetaData()	
        Dim strSQL
		strSQL = "DELETE * FROM tblMetaData WHERE MetaId = " & Me.ID
        RunSQL strSQL		
		DeleteMetaData = eval(Not Me.IsError)
	End Function
	
	Public Function GetMetaDataById()
        Dim strSQL
        strSQL = "SELECT m.MetaId, m.WebPage, p.WebPageName, m.MetaKeywords, m.MetaDescription"
		strSQL = strSQL & " FROM tblMetaData m, tblWebPages p WHERE m.WebPage = p.PageId AND m.MetaId = " & Me.ID
        FillObjectFromRS(strSQL)
	End Function
	
	Public Function GetMetaDataByPageId()
        Dim strSQL
        strSQL = "SELECT m.MetaId, m.WebPage, p.WebPageName, m.MetaKeywords, m.MetaDescription"
		strSQL = strSQL & " FROM tblMetaData m, tblWebPages p WHERE m.WebPage = p.PageId AND p.PageId = " & Me.WebPageId
        FillObjectFromRS(strSQL)
	End Function
		
	Public Function GetMetaData()
        Dim strSQL
        strSQL = "SELECT m.MetaId, m.WebPage, p.WebPageName, m.MetaKeywords, m.MetaDescription"
		strSQL = strSQL & " FROM tblMetaData m, tblWebPages p WHERE m.WebPage = p.PageId"
		FillObjectFromRS(strSQL)
	End Function
		
	'***********************************************
	' PRIVATE METHODS
	'***********************************************
    Private Function FillObjectFromRS(p_strSQL)

        Dim rs
        Set rs = LoadRSFromDB(p_strSQL)		
		
        Dim oMeta
		Dim counter
		counter = 0
		Do While Not rs.EOF
			counter = counter + 1
			Set oMeta = New cMetaData
			oMeta.ID = CInt(rs("MetaId"))
			oMeta.WebPageId = rs("WebPage")
			oMeta.WebPage = rs("WebPageName")
			oMeta.MetaKeywords = QuoteCleanup(rs("MetaKeywords"))
			oMeta.MetaDescription = QuoteCleanup(rs("MetaDescription"))
			m_MetaData.Add oMeta.ID, oMeta
            rs.MoveNext
        Loop
		Me.MetaCount = counter
		rs.Close
        Set rs = Nothing
    End Function	
		
End Class
%>
<%
Class cPageInfo

	Private m_ReturnCode
	Private m_ID
	Private m_WebPage
	Private m_PageInfo
	Private m_WebPageCount
	
	Sub Class_Initialize()
		Set m_PageInfo = Server.CreateObject("Scripting.Dictionary")
		Me.ID = 0
	End Sub
	
	Sub Class_Terminate()
		Set m_PageInfo = Nothing
	End Sub
	
	Public Property Get IsError
		IsError = eval(Me.ID < 0)
	End Property
		
	Public Property Get PageInfo()
		Set PageInfo = m_PageInfo
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
	
	Public Property Let WebPage(p_Value)
		m_WebPage = p_Value
	End Property
	Public Property Get WebPage
		WebPage = m_WebPage
	End Property

	Public Property Let WebPageCount(p_Value)
		m_WebPageCount = p_Value
	End Property
	Public Property Get WebPageCount()
		WebPageCount = m_WebPageCount
	End Property
	
	Public Function AddPageInfo()		
        If Me.ID < 1 then
            Dim arr1, arr2
            arr1 = Array("WebPageName")
            arr2 = Array(Me.WebPage)
            Me.ID = InsertRecord("tblWebPages", "PageId", arr1, arr2)            
        End if
		AddPageInfo = Eval(Not Me.IsError)
	End Function
		
	Public Function DeletePageInfo()	
        Dim strSQL
		strSQL = "DELETE * FROM tblWebPages WHERE PageId = " & Me.ID
        RunSQL strSQL		
		DeletePageInfo = eval(Not Me.IsError)
	End Function
	
	Public Function GetPageInfoById()
        Dim strSQL
        strSQL = "SELECT PageId, WebPageName FROM tblWebPages WHERE PageId = " & Me.ID
        FillObjectFromRS(strSQL)
	End Function
	
	Public Function GetPageInfo()
        Dim strSQL
        strSQL = "SELECT PageId, WebPageName FROM tblWebPages"
		FillObjectFromRS(strSQL)
	End Function
		
	'***********************************************
	' PRIVATE METHODS
	'***********************************************
    Private Function FillObjectFromRS(p_strSQL)

        Dim rs
        Set rs = LoadRSFromDB(p_strSQL)		
		
        Dim oPageInfo
		Dim counter
		counter = 0
		Do While Not rs.EOF
			counter = counter + 1
			Set oPageInfo = New cPageInfo
			oPageInfo.ID = CInt(rs("PageId"))
			oPageInfo.WebPage = rs("WebPageName")
			m_PageInfo.Add oPageInfo.ID, oPageInfo
            rs.MoveNext
        Loop
		Me.WebPageCount = counter
		rs.Close
        Set rs = Nothing
    End Function	
		
End Class
%>
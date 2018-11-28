<%
Class cGallery

	Private m_ID
	Private m_Galleries
	Private m_GalleryLastName
	Private m_GalleryName
	Private m_ExpirationDate
	Private m_GalleryUser
	Private m_GalleryUserId
	Private m_ActiveFlag
	Private m_GalleryCount
	
	Sub Class_Initialize()
		Set m_Galleries = Server.CreateObject("Scripting.Dictionary")
		Me.ID = 0
	End Sub
	
	Sub Class_Terminate()
		Set m_Galleries = Nothing
	End Sub
	
	Public Property Get IsError
		IsError = eval(Me.ID < 0)
	End Property
		
	Public Property Get Galleries()
		Set Galleries = m_Galleries
	End Property
		
	Public Property Let ID(p_Value)
		m_ID = p_Value
	End Property
	Public Property Get ID()
		ID = CLng(m_ID)
	End Property	

	Public Property Let GalleryName(p_Value)
		m_GalleryName = p_Value
	End Property
	Public Property Get GalleryName()
		GalleryName = m_GalleryName
	End Property
	
	Public Property Let GalleryLastName(p_Value)
		m_GalleryLastName = p_Value
	End Property
	Public Property Get GalleryLastName()
		GalleryLastName = m_GalleryLastName
	End Property

	Public Property Let GalleryUser(p_Value)
		m_GalleryUser = p_Value
	End Property
	Public Property Get GalleryUser()
		GalleryUser = m_GalleryUser
	End Property
		
	Public Property Let GalleryUserId(p_Value)
		m_GalleryUserId = p_Value
	End Property
	Public Property Get GalleryUserId()
		GalleryUserId = m_GalleryUserId
	End Property	

	Public Property Let ExpirationDate(p_Value)
		m_ExpirationDate = p_Value
	End Property
	Public Property Get ExpirationDate()
		ExpirationDate = m_ExpirationDate
	End Property
		
	Public Property Let ActiveFlag(p_Value)
		m_ActiveFlag = p_Value
	End Property
	Public Property Get ActiveFlag()
		ActiveFlag = m_ActiveFlag
	End Property	

	Public Property Let GalleryCount(p_Value)
		m_GalleryCount = p_Value
	End Property
	Public Property Get GalleryCount()
		GalleryCount = CLng(m_GalleryCount)
	End Property
	
	Public Function AddGallery()
        If Me.ID < 1 then					
            Dim arr1, arr2
            arr1 = Array("GalleryLastName", "GalleryName", "ExpirationDate", "ActiveFlag", "UserId")
            arr2 = Array(Me.GalleryLastName,SingleQuotes(Me.GalleryName), Me.ExpirationDate, Me.ActiveFlag, Me.GalleryUserId)
            Me.ID = InsertRecord("tblGallery", "GalleryId", arr1, arr2)            
        End if
		AddGalleryName = Eval(Not Me.IsError)
	End Function
	
	Public Function UpdateGallery()
	    Dim strSQL
		strSQL = strSQL & " UPDATE tblGallery SET "
		strSQL = strSQL & " GalleryLastName = '" & Me.GalleryLastName & "', "
		strSQL = strSQL & " GalleryName = '" & SingleQuotes(Me.GalleryName) & "',"
		strSQL = strSQL & " ExpirationDate = #" & Me.ExpirationDate & "#,"
		strSQL = strSQL & " UserId = " & Me.GalleryUserId & ","
		strSQL = strSQL & " ActiveFlag = " & Me.ActiveFlag
		strSQL = strSQL & " WHERE GalleryId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function DeleteGallery()
        Dim strSQL
		strSQL = "DELETE * FROM tblGallery WHERE GalleryId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function GetGalleryById()
        Dim strSQL
        strSQL = "SELECT g.GalleryId, g.GalleryLastName, g.GalleryName, g.ExpirationDate, g.ActiveFlag, g.UserId,  u.FullName"
		strSQL = strSQL & " FROM tblGallery g, tblUsers u WHERE g.GalleryId = " & Me.ID & " And u.UserId = g.UserId ORDER BY g.ExpirationDate DESC"
        FillObjectFromRS(strSQL)
	End Function
	
	Public Function GetGalleryByLastName()
        Dim strSQL
        strSQL = "SELECT g.GalleryId, g.GalleryLastName, g.GalleryName, g.ExpirationDate, g.ActiveFlag, g.UserId,  u.FullName"
		strSQL = strSQL & " FROM tblGallery g, tblUsers u WHERE g.GalleryLastName = '" & Me.GalleryLastName & "' And u.UserId = g.UserId And g.ActiveFlag = True"
		
        FillObjectFromRS(strSQL)
	End Function	
	
	Public Function GetGallery()
        Dim strSQL
        strSQL = "SELECT g.GalleryId, g.GalleryLastName, g.GalleryName, g.ExpirationDate, g.ActiveFlag, g.UserId, u.FullName"
		strSQL = strSQL & " FROM tblGallery g, tblUsers u WHERE u.UserId = g.UserId ORDER BY g.ExpirationDate DESC"
        FillObjectFromRS(strSQL)
	End Function
	
	Public Function GetActiveGallery()
        Dim strSQL
        strSQL = "SELECT g.GalleryId, g.GalleryLastName, g.GalleryName, g.ExpirationDate, g.ActiveFlag, g.UserId, u.FullName"
		strSQL = strSQL & " FROM tblGallery g, tblUsers u WHERE u.UserId = g.UserId And g.ActiveFlag = True ORDER BY g.ExpirationDate DESC"
        FillObjectFromRS(strSQL)
	End Function
	'***********************************************
	' PRIVATE METHODS
	'***********************************************
    Private Function FillObjectFromRS(p_strSQL)

        Dim rs
        Set rs = LoadRSFromDB(p_strSQL)		
		
        Dim oGallery
		Dim counter
		counter = 0
		Do While Not rs.EOF
			counter = counter + 1
			Set oGallery = New cGallery		
			oGallery.ID = rs("GalleryId")
			oGallery.GalleryLastName = rs("GalleryLastName")
			oGallery.GalleryName = QuoteCleanup(rs("GalleryName"))
			oGallery.ExpirationDate = rs("ExpirationDate")
			oGallery.ActiveFlag = rs("ActiveFlag")
			oGallery.GalleryUserId = rs("UserId")
			oGallery.GalleryUser = rs("FullName")
			m_Galleries.Add oGallery.ID, oGallery
            rs.MoveNext
        Loop
		Me.GalleryCount = counter
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
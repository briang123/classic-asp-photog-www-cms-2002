<%
Class cCategory
Private m_Photos
	Private m_ID
	Private m_Categories
	Private m_CategoryText
	Private m_PageOrder
	Private m_CategoryCaption
	Private m_ActiveFlag
	Private m_CategoryCount
	
	Sub Class_Initialize()
		Set m_Categories = Server.CreateObject("Scripting.Dictionary")
		Set m_Photos = Server.CreateObject("Scripting.Dictionary")
		Me.ID = 0
	End Sub
	
	Sub Class_Terminate()
		Set m_Categories = Nothing
		Set m_Photos = Nothing
	End Sub
	
	Public Property Get IsError
		IsError = eval(Me.ID < 0)
	End Property
		
	Public Property Get Categories()
		Set Categories = m_Categories
	End Property
		
	Public Property Get Photos()
		Set Photos = m_Photos
	End Property

	Public Property Let ID(p_Value)
		m_ID = p_Value
	End Property
	Public Property Get ID()
		ID = CLng(m_ID)
	End Property	

	Public Property Let CategoryText(p_Value)
		m_CategoryText = p_Value
	End Property
	Public Property Get CategoryText()
		CategoryText = m_CategoryText
	End Property
	
	Public Property Let PageOrder(p_Value)
		m_PageOrder = CInt(p_Value)
	End Property
	Public Property Get PageOrder()
		PageOrder = m_PageOrder
	End Property
	
	Public Property Let CategoryCaption(p_Value)
		m_CategoryCaption = p_Value
	End Property
	Public Property Get CategoryCaption()
		CategoryCaption = m_CategoryCaption
	End Property
	
	Public Property Let ActiveFlag(p_Value)
		m_ActiveFlag = p_Value
	End Property
	Public Property Get ActiveFlag()
		ActiveFlag = m_ActiveFlag
	End Property	

	Public Property Let CategoryCount(p_Value)
		m_CategoryCount = p_Value
	End Property
	Public Property Get CategoryCount()
		CategoryCount = CLng(m_CategoryCount)
	End Property
	
	Public Function AddCategoryText()
        If Me.ID < 1 then					
            Dim arr1, arr2
            arr1 = Array("CategoryText", "CategoryCaption", "PageOrder", "ActiveFlag")
            arr2 = Array(SingleQuotes(Me.CategoryText), SingleQuotes(Me.CategoryCaption), Me.PageOrder, Me.ActiveFlag)
            Me.ID = InsertRecord("tblCategory", "CategoryId", arr1, arr2)            
        End if
		AddCategoryText = Eval(Not Me.IsError)
	End Function
	
	Public Function UpdateCategoryText()
	    Dim strSQL
		strSQL = strSQL & " UPDATE tblCategory SET "
		strSQL = strSQL & " CategoryText = '" & SingleQuotes(Me.CategoryText) & "', "
		strSQL = strSQL & " PageOrder = " & Me.PageOrder & ", "
		strSQL = strSQL & " CategoryCaption = '" & SingleQuotes(Me.CategoryCaption) & "', "
		strSQL = strSQL & " ActiveFlag = " & Me.ActiveFlag
		strSQL = strSQL & " WHERE CategoryId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function DeleteCategoryText()
        Dim strSQL
		strSQL = "DELETE * FROM tblCategory WHERE CategoryId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function GetCategoryTextById()
        Dim strSQL
        strSQL = "SELECT CategoryId, CategoryText, CategoryCaption, PageOrder, ActiveFlag FROM tblCategory WHERE CategoryId= " & Me.ID    
        FillObjectFromRS(strSQL)
	End Function
	
	Public Function GetCategoryText()
        Dim strSQL
        strSQL = "SELECT CategoryId, CategoryText, CategoryCaption, PageOrder, ActiveFlag FROM tblCategory ORDER BY PageOrder ASC"
        FillObjectFromRS(strSQL)
	End Function
	
	Public Function GetActiveCategoryText()
        Dim strSQL
        strSQL = "SELECT c.CategoryId, c.CategoryText, c.CategoryCaption, c.PageOrder, c.ActiveFlag FROM tblCategory c, tblCatPhotos p "
		strSQL = strSQL & "WHERE c.CategoryId = p.CategoryId AND c.ActiveFlag = -1 AND p.ActiveFlag = -1 "
		strSQL = strSQL & "GROUP BY c.CategoryId, c.CategoryText, c.CategoryCaption, c.PageOrder, c.ActiveFlag "
		strSQL = strSQL & "HAVING COUNT(*) > 0 ORDER BY c.PageOrder ASC"
        FillObjectFromRS(strSQL)
	End Function
	'***********************************************
	' PRIVATE METHODS
	'***********************************************
    Private Function FillObjectFromRS(p_strSQL)

        Dim rs
        Set rs = LoadRSFromDB(p_strSQL)		
		
        Dim oCategory
		Dim counter
		counter = 0
		Do While Not rs.EOF
			counter = counter + 1
			Set oCategory = New cCategory		
			oCategory.ID = rs("CategoryId")
			oCategory.CategoryText = QuoteCleanup(rs("CategoryText"))
			oCategory.PageOrder = rs("PageOrder")
			oCategory.CategoryCaption = QuoteCleanup(rs("CategoryCaption"))
			oCategory.ActiveFlag = rs("ActiveFlag")
			m_Categories.Add oCategory.ID, oCategory
            rs.MoveNext
        Loop
		Me.CategoryCount = counter
		rs.Close
        Set rs = Nothing
    End Function

	Public Function GetActivePhotosByCategory()
        Dim strSQL
        strSQL = "SELECT p.PhotosId, p.CategoryId, p.LargeImage, p.ThumbImage, p.ImageOrder, p.Caption, p.ActiveFlag"
		strSQL = strSQL & " FROM tblCatPhotos p, tblCategory c WHERE p.CategoryId = " & Me.ID & " AND p.ActiveFlag = -1 AND p.CategoryId = c.CategoryId ORDER BY p.ImageOrder ASC"
        FillPhotosByCategory(strSQL)
	End Function

    Private Function FillPhotosByCategory(p_strSQL)

        Dim rs
        Set rs = LoadRSFromDB(p_strSQL)		
		
        Dim oPhotos
		Dim counter
		counter = 0
		Do While Not rs.EOF
			counter = counter + 1
			Set oPhotos = New cPhotos		
			oPhotos.ID = rs("PhotosId")
			oPhotos.CategoryId = rs("CategoryId")
			oPhotos.LargeImage = rs("LargeImage")
			oPhotos.ThumbImage = rs("ThumbImage")
			oPhotos.Caption = rs("Caption")

			'response.write oPhotos.LargeImage & "<br>"
			
			m_Photos.Add oPhotos.ID, oPhotos
            rs.MoveNext
        Loop
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


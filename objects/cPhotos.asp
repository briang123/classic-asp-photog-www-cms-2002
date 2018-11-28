<%
Class cPhotos

	Private m_ID
	Private m_Photos
	Private m_PhotoType
	Private m_WebPageId
	Private m_GalleryId
	Private m_CategoryId
	Private m_LargeImage
	Private m_ThumbImage
	Private m_Caption
	Private m_ImageOrder
	Private m_ActiveFlag
	Private m_PhotosCount
	Private m_GalleryLastName
		
	Sub Class_Initialize()
		Set m_Photos = Server.CreateObject("Scripting.Dictionary")
		Me.ID = 0
	End Sub
	
	Sub Class_Terminate()
		Set m_Photos = Nothing
	End Sub
	
	Public Property Get IsError
		IsError = eval(Me.ID < 0)
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
	
	Public Property Let WebPageId(p_Value)
		m_WebPageId = p_Value
	End Property
	Public Property Get WebPageId()
		WebPageId = m_WebPageId
	End Property		

	Public Property Let GalleryId(p_Value)
		m_GalleryId = p_Value
	End Property
	Public Property Get GalleryId()
		GalleryId = CLng(m_GalleryId)
	End Property	

	Public Property Let CategoryId(p_Value)
		m_CategoryId = p_Value
	End Property
	Public Property Get CategoryId()
		CategoryId = CLng(m_CategoryId)
	End Property	

	Public Property Let LargeImage(p_Value)
		m_LargeImage = p_Value
	End Property
	Public Property Get LargeImage()
		LargeImage = m_LargeImage
	End Property
	
	Public Property Let ThumbImage(p_Value)
		m_ThumbImage = p_Value
	End Property
	Public Property Get ThumbImage()
		ThumbImage = m_ThumbImage
	End Property
	
	Public Property Let Caption(p_Value)
		m_Caption = p_Value
	End Property
	Public Property Get Caption()
		Caption = m_Caption
	End Property	

	Public Property Let ImageOrder(p_Value)
		m_ImageOrder = p_Value
	End Property
	Public Property Get ImageOrder()
		ImageOrder = m_ImageOrder
	End Property
		
	Public Property Let ActiveFlag(p_Value)
		m_ActiveFlag = p_Value
	End Property
	Public Property Get ActiveFlag()
		ActiveFlag = m_ActiveFlag
	End Property	

	'Possible Types(GALLERY,CATEGORY)
	Public Property Let PhotoType(p_Value)
		m_PhotoType = p_Value
	End Property
	Public Property Get PhotoType()
		PhotoType = UCase(m_PhotoType)
	End Property

	Public Property Let PhotosCount(p_Value)
		m_PhotosCount = p_Value
	End Property
	Public Property Get PhotosCount()
		PhotosCount = CLng(m_PhotosCount)
	End Property

	Public Property Let GalleryLastName(p_Value)
		m_GalleryLastName = p_Value
	End Property
	Public Property Get GalleryLastName()
		GalleryLastName = m_GalleryLastName
	End Property


'******************************************************************************************************************************************************
'W E B S I T E  P H O T O G R A P H  I N F O R M A T I O N
'******************************************************************************************************************************************************

	Public Function AddSitePhoto()
        If Me.ID < 1 then					
            Dim arr1, arr2
            arr1 = Array("PageId", "LargeImage", "ImageOrder", "ActiveFlag")
            arr2 = Array(Me.WebPageId, Me.LargeImage, Me.ImageOrder, Me.ActiveFlag)
            Me.ID = InsertRecord("tblSitePhotos", "PhotosId", arr1, arr2)            
        End if
		AddSitePhoto = Eval(Not Me.IsError)
	End Function

	Public Function UpdateSitePhotoInfo()
	    Dim strSQL
		strSQL = strSQL & " UPDATE tblSitePhotos SET "
		strSQL = strSQL & " PageId = " & Me.WebPageId & ", "
		strSQL = strSQL & " ImageOrder = " & Me.ImageOrder & ","
		strSQL = strSQL & " ActiveFlag = " & Me.ActiveFlag
		strSQL = strSQL & " WHERE PhotosId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function UpdateSitePhotos()
	    Dim strSQL
		strSQL = strSQL & " UPDATE tblSitePhotos SET "
		strSQL = strSQL & " PageId = " & Me.WebPageId & ","
		strSQL = strSQL & " LargeImage = " & SingleQuotes(Me.LargeImage) & ","
		strSQL = strSQL & " ImageOrder = " & Me.ImageOrder & ","
		strSQL = strSQL & " ActiveFlag = " & Me.ActiveFlag
		strSQL = strSQL & " WHERE PhotosId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function DeleteSitePhotos()
        Dim strSQL
		strSQL = "DELETE * FROM tblSitePhotos WHERE PhotosId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
			
	Public Function GetSitePhotosById()
        Dim strSQL
        strSQL = "SELECT p.PhotosId, p.PageId, s.WebPageName, p.LargeImage, p.ImageOrder, p.ActiveFlag"
		strSQL = strSQL & " FROM tblSitePhotos p, tblWebPages s WHERE s.PageId = p.WebPageId AND p.PhotosId = " & Me.ID
        FillSiteObjectFromRS(strSQL)
	End Function
	
	Public Function GetSitePhotos()
        Dim strSQL
        strSQL = "SELECT p.PhotosId, p.PageId, p.LargeImage, p.ImageOrder, p.ActiveFlag"
		strSQL = strSQL & " FROM tblSitePhotos p ORDER BY p.ActiveFlag, p.PageId, p.ImageOrder ASC"
        FillSiteObjectFromRS(strSQL)
	End Function

	Public Function GetSitePhotosByPage()
        Dim strSQL
        strSQL = "SELECT p.PhotosId, p.PageId, s.WebPageName, p.LargeImage, p.ImageOrder, p.ActiveFlag"
		strSQL = strSQL & " FROM tblSitePhotos p, tblWebPages s WHERE s.PageId = p.PageId AND s.PageId = " & Me.WebPageId & " ORDER BY p.ActiveFlag, p.ImageOrder ASC"
        FillSiteObjectFromRS(strSQL)
	End Function
	
	Public Function GetActiveSitePhotosByPage()
        Dim strSQL
        strSQL = "SELECT p.PhotosId, p.PageId, s.WebPageName, p.LargeImage, p.ImageOrder, p.ActiveFlag"
		strSQL = strSQL & " FROM tblSitePhotos p, tblWebPages s WHERE s.PageId = p.PageId AND s.PageId = " & Me.WebPageId & " AND p.ActiveFlag = -1 ORDER BY p.ImageOrder ASC"
        FillSiteObjectFromRS(strSQL)
	End Function
		
    Private Function FillSiteObjectFromRS(p_strSQL)

        Dim rs
        Set rs = LoadRSFromDB(p_strSQL)		
		
        Dim oPhotos
		Dim counter
		counter = 0
		Do While Not rs.EOF
			counter = counter + 1
			Set oPhotos = New cPhotos		
			oPhotos.ID = rs("PhotosId")
			oPhotos.WebPageId = rs("PageId")
			oPhotos.LargeImage = rs("LargeImage")
			oPhotos.ImageOrder = rs("ImageOrder")
			oPhotos.ActiveFlag = rs("ActiveFlag")
			m_Photos.Add oPhotos.ID, oPhotos
            rs.MoveNext
        Loop
		Me.PhotosCount = counter
		rs.Close
        Set rs = Nothing
    End Function
	
'******************************************************************************************************************************************************
'G A L L E R Y  I N F O R M A T I O N
'******************************************************************************************************************************************************


	Public Function AddGalleryPhoto()
        If Me.ID < 1 then					
            Dim arr1, arr2
            arr1 = Array("GalleryId", "LargeImage", "ThumbImage", "Caption", "ImageOrder", "ActiveFlag")
            arr2 = Array(Me.GalleryId, Me.LargeImage, Me.ThumbImage, Me.Caption, Me.ImageOrder, Me.ActiveFlag)
            Me.ID = InsertRecord("tblGalPhotos", "PhotosId", arr1, arr2)            
        End if
		AddGalleryPhoto = Eval(Not Me.IsError)
	End Function
	
	Public Function UpdateGalleryPhotoInfo()
	    Dim strSQL
		strSQL = strSQL & " UPDATE tblGalPhotos SET "
		strSQL = strSQL & " Caption = '" & SingleQuotes(Me.Caption) & "'," 
		strSQL = strSQL & " ImageOrder = " & Me.ImageOrder & ","
		strSQL = strSQL & " ActiveFlag = " & Me.ActiveFlag
		strSQL = strSQL & " WHERE PhotosId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function UpdateGalleryPhotos()
	    Dim strSQL
		strSQL = strSQL & " UPDATE tblGalPhotos SET "
		strSQL = strSQL & " GalleryId = " & Me.GalleryId & ","
		strSQL = strSQL & " LargeImage = " & SingleQuotes(Me.LargeImage) & ","
		strSQL = strSQL & " ThumbImage = " & SingleQuotes(Me.ThumbImage) & ","
		strSQL = strSQL & " Caption = " & SingleQuotes(Me.Caption) & "," 
		strSQL = strSQL & " ImageOrder = " & Me.ImageOrder & ","
		strSQL = strSQL & " ActiveFlag = " & Me.ActiveFlag
		strSQL = strSQL & " WHERE PhotosId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function DeleteGalleryPhotos()
        Dim strSQL
		strSQL = "DELETE * FROM tblGalPhotos WHERE PhotosId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function DeletePhotosByGallery()
        Dim strSQL
		strSQL = "DELETE * FROM tblGalPhotos WHERE GalleryId = " & Me.GalleryId
        Call RunSQL(strSQL)
	End Function
		
	Public Function GetGalleryPhotosById()
        Dim strSQL
        strSQL = "SELECT p.PhotosId, p.GalleryId, g.GalleryName, p.LargeImage, p.ThumbImage, p.Caption, p.ImageOrder, p.ActiveFlag"
		strSQL = strSQL & " FROM tblGalPhotos p, tblGallery g WHERE g.GalleryId = p.GalleryId AND p.PhotosId = " & Me.ID
        FillObjectFromRS(strSQL)
	End Function
	
	Public Function GetGalleryPhotos()
        Dim strSQL
        strSQL = "SELECT p.PhotosId, p.GalleryId, g.GalleryName, p.LargeImage, p.ThumbImage, p.Caption, p.ImageOrder, p.ActiveFlag"
		strSQL = strSQL & " FROM tblGalPhotos p, tblGallery g WHERE p.GalleryId = g.GalleryId ORDER BY p.ImageOrder ASC"
        FillObjectFromRS(strSQL)
	End Function

	Public Function GetPhotosByGallery()
        Dim strSQL
        strSQL = "SELECT p.PhotosId, p.GalleryId, p.LargeImage, p.ThumbImage, p.ImageOrder, p.Caption, p.ActiveFlag"
		strSQL = strSQL & " FROM tblGalPhotos p, tblGallery g WHERE p.GalleryId = " & Me.GalleryId & " AND p.GalleryId = g.GalleryId ORDER BY p.ImageOrder ASC"
        FillObjectFromRS(strSQL)
	End Function
	
	Public Function GetPhotosByGalleryLastName()
        Dim strSQL
        strSQL = "SELECT p.PhotosId, p.GalleryId, p.LargeImage, p.ThumbImage, p.ImageOrder, p.Caption, p.ActiveFlag"
		strSQL = strSQL & " FROM tblGalPhotos p, tblGallery g WHERE g.GalleryLastName = '" & SingleQuotes(LCase(Me.GalleryLastName)) & "' AND p.ActiveFlag = True AND p.GalleryId = g.GalleryId ORDER BY p.ImageOrder ASC"
        FillObjectFromRS(strSQL)
	End Function
	
'******************************************************************************************************************************************************
'C A T E G O R Y  I N F O R M A T I O N
'******************************************************************************************************************************************************

	Public Function AddCategoryPhoto()
        If Me.ID < 1 then					
            Dim arr1, arr2
            arr1 = Array("CategoryId", "LargeImage", "ThumbImage", "Caption", "ImageOrder", "ActiveFlag")
            arr2 = Array(Me.CategoryId, Me.LargeImage, Me.ThumbImage, Me.Caption, Me.ImageOrder, Me.ActiveFlag)
            Me.ID = InsertRecord("tblCatPhotos", "PhotosId", arr1, arr2)            
        End if
		AddCategoryPhoto = Eval(Not Me.IsError)
	End Function
	
	Public Function UpdateCategoryPhotoInfo()
	    Dim strSQL
		strSQL = strSQL & " UPDATE tblCatPhotos SET "
		strSQL = strSQL & " Caption = '" & SingleQuotes(Me.Caption) & "'," 
		strSQL = strSQL & " ImageOrder = " & Me.ImageOrder & ","
		strSQL = strSQL & " ActiveFlag = " & Me.ActiveFlag
		strSQL = strSQL & " WHERE PhotosId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function UpdateCategoryPhotos()
	    Dim strSQL
		strSQL = strSQL & " UPDATE tblCatPhotos SET "
		strSQL = strSQL & " CategoryId = " & Me.CategoryId & ","
		strSQL = strSQL & " LargeImage = " & SingleQuotes(Me.LargeImage) & ","
		strSQL = strSQL & " ThumbImage = " & SingleQuotes(Me.ThumbImage) & ","
		strSQL = strSQL & " Caption = " & SingleQuotes(Me.Caption) & "," 
		strSQL = strSQL & " ImageOrder = " & Me.ImageOrder & ","
		strSQL = strSQL & " ActiveFlag = " & Me.ActiveFlag
		strSQL = strSQL & " WHERE PhotosId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function DeleteCategoryPhotos()
        Dim strSQL
		strSQL = "DELETE * FROM tblCatPhotos WHERE PhotosId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function DeletePhotosByCategory()
        Dim strSQL
		strSQL = "DELETE * FROM tblCatPhotos WHERE CategoryId = " & Me.CategoryId
        Call RunSQL(strSQL)
	End Function
		
	Public Function GetCategoryPhotosById()
        Dim strSQL
        strSQL = "SELECT p.PhotosId, p.CategoryId, c.CategoryText, p.LargeImage, p.ThumbImage, p.Caption, p.ImageOrder, p.ActiveFlag"
		strSQL = strSQL & " FROM tblCatPhotos p, tblCategory c WHERE c.CategoryId = p.CategoryId AND p.PhotosId = " & Me.ID
        FillObjectFromRS(strSQL)
	End Function
	
	Public Function GetCategoryPhotos()
        Dim strSQL
        strSQL = "SELECT p.PhotosId, p.CategoryId, c.CategoryText, p.LargeImage, p.ThumbImage, p.Caption, p.ImageOrder, p.ActiveFlag"
		strSQL = strSQL & " FROM tblCatPhotos p, tblCategory c WHERE p.CategoryId = c.CategoryId ORDER BY c.PageOrder ASC"
        FillObjectFromRS(strSQL)
	End Function

	Public Function GetPhotosByCategory()
        Dim strSQL
        strSQL = "SELECT p.PhotosId, p.CategoryId, p.LargeImage, p.ThumbImage, p.ImageOrder, p.Caption, p.ActiveFlag"
		strSQL = strSQL & " FROM tblCatPhotos p, tblCategory c WHERE p.CategoryId = " & Me.CategoryId & " AND p.CategoryId = c.CategoryId ORDER BY p.ImageOrder ASC"
        FillObjectFromRS(strSQL)
	End Function

	Public Function GetActivePhotosByCategory()
        Dim strSQL
        strSQL = "SELECT p.PhotosId, p.CategoryId, p.LargeImage, p.ThumbImage, p.ImageOrder, p.Caption, p.ActiveFlag"
		strSQL = strSQL & " FROM tblCatPhotos p, tblCategory c WHERE p.CategoryId = " & Me.CategoryId & " AND p.ActiveFlag = -1 AND p.CategoryId = c.CategoryId ORDER BY p.ImageOrder ASC"
        FillObjectFromRS(strSQL)
	End Function


	'***********************************************
	' PRIVATE METHODS
	'***********************************************
    Private Function FillObjectFromRS(p_strSQL)

        Dim rs
        Set rs = LoadRSFromDB(p_strSQL)		
		
        Dim oPhotos
		Dim counter
		counter = 0
		Do While Not rs.EOF
			counter = counter + 1
			Set oPhotos = New cPhotos		
			oPhotos.ID = rs("PhotosId")
			
			If Me.PhotoType = "GALLERY" Then
				oPhotos.GalleryId = rs("GalleryId")
				oPhotos.CategoryId = 0
			ElseIf Me.PhotoType = "CATEGORY" Then
				oPhotos.CategoryId = rs("CategoryId")
				oPhotos.GalleryId = 0
			End If
			oPhotos.LargeImage = rs("LargeImage")
			oPhotos.ThumbImage = rs("ThumbImage")
			oPhotos.Caption = rs("Caption")
			oPhotos.ImageOrder = rs("ImageOrder")
			oPhotos.ActiveFlag = rs("ActiveFlag")
			m_Photos.Add oPhotos.ID, oPhotos
            rs.MoveNext
        Loop
		Me.PhotosCount = counter
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


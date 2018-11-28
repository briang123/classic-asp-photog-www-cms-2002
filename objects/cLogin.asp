<%
Class cLogin

	Private m_ID
	Private m_Login
	Private m_Password
	Private m_Email
	Private m_FullName
	Private m_Phone
	Private m_Address1
	Private m_Address2
	Private m_City
	Private m_StateCode
	Private m_Zip
	Private m_Comments
	Private m_Photographer
	Private m_Expire
	Private m_IsAdmin
	Private m_GalleryName
	Private m_LoginCount
	Private m_Logins
	
	Sub Class_Initialize()
		Set m_Logins = Server.CreateObject("Scripting.Dictionary")
		Me.ID = 0
	End Sub
	
	Sub Class_Terminate()
		Set m_Logins = Nothing
	End Sub
	
	Public Property Get IsError
		IsError = eval(Me.ID < 0)
	End Property
		
	Public Property Get Logins()
		Set Logins = m_Logins
	End Property
		
	Public Property Let ID(p_Value)
		m_ID = p_Value
	End Property
	Public Property Get ID()
		ID = CLng(m_ID)
	End Property	

	Public Property Let Login(p_Value)
		m_Login = p_Value
	End Property
	Public Property Get Login()
		Login = m_Login
	End Property
	
	Public Property Let Password(p_Value)
		m_Password = p_Value
	End Property
	Public Property Get Password()
		Password = m_Password
	End Property
		
	Public Property Let Email(p_Value)
		m_Email = p_Value
	End Property
	Public Property Get Email()
		Email = m_Email
	End Property	
	
	Public Property Let FullName(p_Value)
		m_FullName = p_Value
	End Property
	Public Property Get FullName()
		FullName = m_FullName
	End Property	
	
	Public Property Let Phone(p_Value)
		m_Phone = p_Value
	End Property
	Public Property Get Phone()
		Phone = m_Phone
	End Property

	Public Property Let Address1(p_Value)
		m_Address1 = p_Value
	End Property
	Public Property Get Address1()
		Address1 = m_Address1
	End Property

	Public Property Let Address2(p_Value)
		m_Address2 = p_Value
	End Property
	Public Property Get Address2()
		Address2 = m_Address2
	End Property
	
	Public Property Let City(p_Value)
		m_City = p_Value
	End Property
	Public Property Get City()
		City = m_City
	End Property

	Public Property Let StateCode(p_Value)
		m_StateCode = p_Value
	End Property
	Public Property Get StateCode()
		StateCode = m_StateCode
	End Property

	Public Property Let Zip(p_Value)
		m_Zip = p_Value
	End Property
	Public Property Get Zip()
		Zip = m_Zip
	End Property

	Public Property Let Comments(p_Value)
		m_Comments = p_Value
	End Property
	Public Property Get Comments()
		Comments = m_Comments
	End Property

	Public Property Let Photographer(p_Value)
		m_Photographer = p_Value
	End Property
	Public Property Get Photographer()
		Photographer = m_Photographer
	End Property

	Public Property Let Expire(p_Value)
		m_Expire = p_Value
	End Property
	Public Property Get Expire()
		Expire = m_Expire
	End Property

	Public Property Let IsAdmin(p_Value)
		m_IsAdmin = p_Value
	End Property
	Public Property Get IsAdmin()
		IsAdmin = m_IsAdmin
	End Property
	
	Public Property Let GalleryName(p_Value)
		m_GalleryName = p_Value
	End Property
	Public Property Get GalleryName()
		GalleryName = m_GalleryName
	End Property
	
	Public Property Let LoginCount(p_Value)
		m_LoginCount = p_Value
	End Property
	Public Property Get LoginCount()
		LoginCount = CLng(m_LoginCount)
	End Property
	
	Public Function AddUser()
        If Me.ID < 1 then
            Dim arr1, arr2
            arr1 = Array("Login","Pwd","Email","FullName","Phone","Address1","Address2","City","StateCode","Zip","Comments","Photographer","Expire","IsAdmin")
            arr2 = Array(Me.Login,Me.Password,Me.Email,Me.FullName,Me.Phone,Me.Address1,Me.Address2,Me.City,Me.StateCode,Me.Zip,Me.Comments,Me.Photographer,Me.Expire,Me.IsAdmin)
            Me.ID = InsertRecord("tblUsers", "UserId", arr1, arr2)            
        End if
		AddUser = Eval(Not Me.IsError)
	End Function
	
	Public Function UpdateUser()
	    Dim strSQL
		strSQL = strSQL & " UPDATE tblUsers SET "
		strSQL = strSQL & " Login = '" & Me.Login & "'," 
		strSQL = strSQL & " Pwd = '" & Me.Password & "'," 
		strSQL = strSQL & " Email = '" & SingleQuotes(Me.Email) & "'," 
		strSQL = strSQL & " FullName = '" & SingleQuotes(Me.FullName) & "'," 
		strSQL = strSQL & " Phone = '" & SingleQuotes(Me.Phone) & "'," 
		strSQL = strSQL & " Address1 = '" & SingleQuotes(Me.Address1) & "',"
		strSQL = strSQL & " Address2 = '" & SingleQuotes(Me.Address2) & "',"
		strSQL = strSQL & " City = '" & SingleQuotes(Me.City) & "',"
		strSQL = strSQL & " StateCode = '" & Me.StateCode & "',"
		strSQL = strSQL & " Zip = '" & Me.Zip & "',"
		strSQL = strSQL & " Comments = '" & SingleQuotes(Me.Comments) & "',"
		strSQL = strSQL & " Photographer = '" & SingleQuotes(Me.Photographer) & "',"
		strSQL = strSQL & " Expire = #" & Me.Expire & "#,"
		strSQL = strSQL & " IsAdmin = " & Me.IsAdmin		
		strSQL = strSQL & " WHERE UserId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function DeleteUser()
        Dim strSQL
		strSQL = "DELETE * FROM tblUsers WHERE UserId = " & Me.ID
        Call RunSQL(strSQL)
	End Function
	
	Public Function GetUserById()
        Dim strSQL
        strSQL = "SELECT UserId, Login, Pwd, Email, FullName, Phone, Address1, Address2, City, StateCode, Zip, Comments, Photographer, Expire, IsAdmin FROM tblUsers WHERE UserId = " & Me.ID    
        Call FillObjectFromRS(strSQL,False)
	End Function
		
	Public Function GetUserByLogin()
        Dim strSQL
		strSQL = "SELECT u.UserId, u.Login, u.Pwd, u.Email, u.FullName, u.Phone, u.Address1, u.Address2, u.City, u.StateCode, u.Zip, u.Comments, u.Photographer, u.Expire, u.IsAdmin, g.GalleryLastName FROM tblUsers u LEFT JOIN tblGallery g ON g.UserId = u.UserId WHERE u.Login = '" & Replace(Me.Login,"'","") & "' AND u.Pwd = '" & Replace(Me.Password,"'","") & "'"
        Call FillObjectFromRS(strSQL,True)
	End Function
	
	Public Function GetUsers()
        Dim strSQL
		strSQL = "SELECT UserId, Login, Pwd, Email, FullName, Phone, Address1, Address2, City, StateCode, Zip, Comments, Photographer, Expire, IsAdmin FROM tblUsers ORDER BY Expire DESC"
        Call FillObjectFromRS(strSQL,False)
	End Function
	
	'***********************************************
	' PRIVATE METHODS
	'***********************************************
    Private Function FillObjectFromRS(p_strSQL,hasGallery)

        Dim rs
        Set rs = LoadRSFromDB(p_strSQL)		
		
        Dim oLogin
		Dim counter
		counter = 0
		Do While Not rs.EOF
			counter = counter + 1
			Set oLogin = New cLogin
			oLogin.ID = rs("UserId")
			oLogin.Login = rs("Login")
			oLogin.Password = rs("Pwd")
			oLogin.Email = QuoteCleanup(rs("Email"))
			oLogin.FullName = QuoteCleanup(rs("FullName"))
			oLogin.Phone = rs("Phone")
			oLogin.Address1 = QuoteCleanup(rs("Address1"))
			oLogin.Address2 = QuoteCleanup(rs("Address2"))
			oLogin.City = QuoteCleanup(rs("City"))
			oLogin.StateCode = rs("StateCode")
			oLogin.Zip = rs("Zip")
			oLogin.Comments = QuoteCleanup(rs("Comments"))
			oLogin.Photographer = QuoteCleanup(rs("Photographer"))
			oLogin.Expire = rs("Expire")
			oLogin.IsAdmin = rs("IsAdmin")
			If hasGallery Then
				oLogin.GalleryName = LCase(rs("GalleryLastName"))
			Else
				oLogin.GalleryName = ""
			End If
			m_Logins.Add counter, oLogin
            rs.MoveNext
        Loop
		Me.LoginCount = counter
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
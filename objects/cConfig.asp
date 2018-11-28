<%
Class cConfig

	Private m_ReturnCode
	Private m_ID
	Private m_ConfigKey
	Private m_ConfigValue
	Private m_ConfigDesc
	Private m_Configs
	Private m_ConfigCount
	
	Sub Class_Initialize()
		Set m_Configs = Server.CreateObject("Scripting.Dictionary")
		Me.ID = 0
	End Sub
	
	Sub Class_Terminate()
		Set m_Configs = Nothing
	End Sub
	
	Public Property Get IsError
		IsError = eval(Me.ID < 0)
	End Property
		
	Public Property Get Configs()
		Set Configs = m_Configs
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
	
	Public Property Let ConfigKey(p_Value)
		m_ConfigKey = p_Value
	End Property
	Public Property Get ConfigKey
		ConfigKey = CStr(m_ConfigKey)
	End Property

	Public Property Let ConfigValue(p_Value)
		m_ConfigValue = p_Value
	End Property
	Public Property Get ConfigValue
		ConfigValue = m_ConfigValue
	End Property

	Public Property Let ConfigDesc(p_Value)
		m_ConfigDesc = p_Value
	End Property
	Public Property Get ConfigDesc
		ConfigDesc = m_ConfigDesc
	End Property

	Public Property Let ConfigCount(p_Value)
		m_ConfigCount = p_Value
	End Property
	Public Property Get ConfigCount()
		ConfigCount = m_ConfigCount
	End Property
	
	Public Function AddConfig()
		If Me.ConfigValue = "" Then 
			Me.ConfigValue = Null
		Else
			Me.ConfigValue = SingleQuotes(Me.ConfigValue)
		End If
		
		If Me.ConfigDesc = "" Then
			Me.ConfigDesc = Null
		Else
			Me.ConfigDesc = SingleQuotes(Me.ConfigDesc)
		End If
		
        If Me.ID < 1 then
            Dim arr1, arr2
            arr1 = Array("ConfigKey", "ConfigValue", "ConfigDesc")
            arr2 = Array(Me.ConfigKey, Me.ConfigValue, Me.ConfigDesc)
            Me.ID = InsertRecord("tblConfig", "ConfigId", arr1, arr2)            
        End if
		AddConfig = Eval(Not Me.IsError)
		Call GetConfigs()		
	End Function
	
	Public Function UpdateConfig()
		If Me.ConfigValue = "" Then 
			Me.ConfigValue = Null
		Else
			Me.ConfigValue = SingleQuotes(Me.ConfigValue)
		End If
		
		If Me.ConfigDesc = "" Then
			Me.ConfigDesc = Null
		Else
			Me.ConfigDesc = SingleQuotes(Me.ConfigDesc)
		End If
		
	    Dim strSQL
		strSQL = strSQL & " UPDATE tblConfig SET "
		strSQL = strSQL & " ConfigKey = '" & SingleQuotes(Me.ConfigKey) & "',"
		strSQL = strSQL & " ConfigValue = '" & Me.ConfigValue & "', "
		strSQL = strSQL & " ConfigDesc = '" & Me.ConfigDesc & "' " 
		strSQL = strSQL & " WHERE ConfigId = " & Me.ID
		RunSQL strSQL            
		
		UpdateConfig = eval(Not Me.IsError)
		Call GetConfigs()		
	End Function
	
	Public Function DeleteConfig()	
        Dim strSQL
		strSQL = "DELETE * FROM tblConfig WHERE ConfigId = " & Me.ID
        RunSQL strSQL		
		DeleteConfig = eval(Not Me.IsError)
	End Function
	
	Public Function GetConfigInfoById()
        Dim strSQL
        strSQL = "SELECT ConfigId, ConfigKey, ConfigValue, ConfigDesc FROM tblConfig WHERE ConfigId = " & Me.ID
        FillObjectFromRS(strSQL)
	End Function
	
	Public Function GetConfigs()
        Dim strSQL
        strSQL = "SELECT ConfigId, ConfigKey, ConfigValue, ConfigDesc FROM tblConfig ORDER BY ConfigKey ASC"
		FillObjectFromRS(strSQL)
	End Function
		
	'***********************************************
	' PRIVATE METHODS
	'***********************************************
    Private Function FillObjectFromRS(p_strSQL)

        Dim rs
        Set rs = LoadRSFromDB(p_strSQL)		
		
        Dim oConfig
		Dim counter
		counter = 0
		Do While Not rs.EOF
			counter = counter + 1
			Set oConfig = New cConfig
			oConfig.ID = CInt(rs("ConfigId"))
			oConfig.ConfigKey = QuoteCleanup(rs("ConfigKey"))
			oConfig.ConfigValue = QuoteCleanup(rs("ConfigValue"))
			oConfig.ConfigDesc = QuoteCleanup(rs("ConfigDesc"))
			m_Configs.Add oConfig.ID, oConfig
			On Error Resume Next
				Application.Lock()
				Call AddAppVariable(oConfig.ConfigKey, QuoteCleanup(oConfig.ConfigValue))
				Application.UnLock()		
			On Error Goto 0			
            rs.MoveNext
        Loop
		Me.ConfigCount = counter
		rs.Close
        Set rs = Nothing
    End Function	
		
End Class
%>
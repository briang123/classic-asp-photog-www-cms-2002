<%
Class cProduct

	Private m_ReturnCode
	Private m_ID
	Private m_MSDSFileName
	Private m_TDSFileName
	Private m_BulletinFileName
	Private m_ProductName
	Private m_LambentFileType
	Private m_UploadDate
	Private m_UploadBy
	Private m_LastModifiedDate
	Private m_LastModifiedBy
	Private m_ActiveFlag
	Private m_Product
	Private m_ProductCount
	
	Sub Class_Initialize()
		Set m_Product = Server.CreateObject("Scripting.Dictionary")
		Me.ID = 0
		Me.MSDSFileName = ""
		Me.TDSFileName = ""
		Me.BulletinFileName = ""
	End Sub
	
	Sub Class_Terminate()
		Set m_Product = Nothing
	End Sub
	
	Public Property Get IsError
		IsError = eval(Me.ReturnCode = 0)
	End Property
		
	Public Property Get Product()
		Set Product = m_Product
	End Property
		
	Public Property Get ReturnCode()
		ReturnCode = m_ReturnCode
	End Property
	
	Public Property Let ID(p_Value)
		m_ID = p_Value
	End Property
	Public Property Get ID()
		ID = CInt(m_ID)
	End Property
	
	Public Property Let ProductName(p_Value)
		m_ProductName = p_Value
	End Property
	Public Property Get ProductName
		ProductName = CStr(m_ProductName)
	End Property	
	
	Public Property Let ActiveFlag(p_Value)
		m_ActiveFlag = p_Value
	End Property
	Public Property Get ActiveFlag()
		ActiveFlag = m_ActiveFlag
	End Property	

	Public Property Let ProductCount(p_Value)
		m_ProductCount = p_Value
	End Property
	Public Property Get ProductCount()
		ProductCount = m_ProductCount
	End Property
	
	Public Function AddProduct()
		Dim cmd
		Set cmd = Server.CreateObject("ADODB.Command")	
		With cmd
			.ActiveConnection = CONNECTION_STRING
			.CommandText = "sp__AddProduct"
			.CommandType = adCmdStoredProc	
			.Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)				
			.Parameters.Append .CreateParameter("@ProductName",adVarChar,adParamInput,100,SingleQuotes(Me.ProductName))		
			.Parameters.Append .CreateParameter("@ActiveFlag",adTinyInt,adParamInput, ,Me.ActiveFlag)
			.Parameters.Append .CreateParameter("@ProductId",adInteger, adParamOutput, ,Me.ID)						
			.Execute , ,adExecuteNoRecords
			m_ID = .Parameters("@ProductId")
			m_ReturnCode = .Parameters("Return")
		End With
		CloseCmd(cmd)
	End Function
	
	Public Function UpdateProduct()
		Dim cmd
		Set cmd = Server.CreateObject("ADODB.Command")	
		With cmd
			.ActiveConnection = CONNECTION_STRING
			.CommandText = "sp__UpdateProduct"
			.CommandType = adCmdStoredProc
			.Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
			.Parameters.Append .CreateParameter("@ProductId",adInteger, adParamInput, ,Me.ID)							
			.Parameters.Append .CreateParameter("@ProductName",adVarChar,adParamInput,100,SingleQuotes(Me.ProductName))		
			.Parameters.Append .CreateParameter("@ActiveFlag",adTinyInt,adParamInput, ,Me.ActiveFlag)
			.Execute , ,adExecuteNoRecords
			m_ReturnCode = .Parameters("Return")
		End With
		CloseCmd(cmd)
	End Function
	
	Public Function DeleteProduct()	
		Dim cmd
		Set cmd = Server.CreateObject("ADODB.Command")	
		With cmd
			.ActiveConnection = CONNECTION_STRING
			.CommandText = "sp__DeleteProduct"
			.CommandType = adCmdStoredProc
			.Parameters.Append .CreateParameter("Return", adInteger, adParamReturnValue)
			.Parameters.Append .CreateParameter("@ProductId", adInteger, adParamInput, ,Me.ID)
			.Execute , ,adExecuteNoRecords
			m_ReturnCode = .Parameters("Return")
		End With
		CloseCmd(cmd)
	End Function
	
	Public Function GetProductInfoById()
		Dim cmd, rs
		Set cmd = Server.CreateObject("ADODB.Command")	
		Set rs = Server.CreateObject("ADODB.RecordSet")
		With cmd
			.ActiveConnection = CONNECTION_STRING
			.CommandText = "sp__GetProductInfoById"
			.CommandType = adCmdStoredProc	
			.Parameters.Append .CreateParameter("@ProductId", adInteger, adParamInput, ,Me.ID)				
			Set rs = .Execute
		End With
		FillObjectFromRS(rs)
		CloseRs(rs)		
		CloseCmd(cmd)	
	End Function
	
	Public Function GetProduct()
		Dim cmd, rs
		Set cmd = Server.CreateObject("ADODB.Command")	
		Set rs = Server.CreateObject("ADODB.RecordSet")
		With cmd
			.ActiveConnection = CONNECTION_STRING
			.CommandText = "sp__GetProduct"
			.CommandType = adCmdStoredProc		
			Set rs = .Execute
		End With
		FillObjectFromRS(rs)
		CloseRs(rs)		
		CloseCmd(cmd)
	End Function
	
	'***********************************************
	' PRIVATE METHODS
	'***********************************************
    Private Function FillObjectFromRS(p_RS)
        Dim oProduct
		Dim counter
		counter = 0
		Do While Not p_RS.EOF
			counter = counter + 1		
            Set oProduct = New cProduct
            oProduct.ID = p_RS("ProductId")
            oProduct.ProductName = QuoteCleanup(p_RS("ProductName"))
			oProduct.ActiveFlag = p_RS("ActiveFlag")
            m_Product.Add oProduct.ID, oProduct
            p_RS.MoveNext
        Loop
		Me.ProductCount = counter
    End Function

End Class
%>
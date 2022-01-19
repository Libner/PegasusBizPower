<%	'on error resume next
	dim vDB, vField, vID, rsBLOB, SQLString
	  
	vDB=Request("DB") 
	If trim(vDB) = "Element" Or trim(vDB) = "Template_Element" Then
		vFieldId = "Element_Id"
	ElseIf trim(vDB) = "Page" Then
		vFieldId = "Page_Id"
	ElseIf trim(vDB) = "Template"Then
		vFieldId = "Template_Id"	
	ElseIf trim(vDB) = "ORGANIZATION"Then
		vFieldId = "ORGANIZATION_Id"	
	Else
		vFieldId = "Product_Id" 	
	End If		
	vField=Request("Field") 
	vID=trim(Request("ID"))
	If isNumeric(vID) Then
	Response.Expires = -1
	Response.Buffer = TRUE
	Response.Clear
	Response.ContentType="image/gif"

	Set con = Server.CreateObject("ADODB.Connection")   
    con.CommandTimeout = 0
    con.ConnectionTimeout = 0
    con.Open Application("ConnectionString")
   
     SQLString = "SELECT " & vField & " FROM " & vDb & "s WHERE " & vFieldId & "=" & vID
     'Response.Write  SQLString
     'Response.End

Set rsBlob = con.Execute(SQLString)
If rsBlob(vField).ActualSize > 0 Then
Response.BinaryWrite rsBlob(vField).getChunk(rsBlob(vField).ActualSize)
End If
Response.End
End If
con.Close()%>
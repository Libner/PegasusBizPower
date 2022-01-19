<%
Function IsEmailValid(strEmail) 

		strEmail = trim(strEmail)
   
		Dim strArray 
		Dim strItem 
		Dim i 
		Dim c 
		Dim blnIsItValid 
	   
		' assume the email address is correct   
		blnIsItValid = True 
		
		If IsNull(strEmail) Then
			blnIsItValid = False 
			IsEmailValid = blnIsItValid 
			Exit Function
		End If
	     
		If InStr(strEmail,"@") < 0 Then
			blnIsItValid = False 
			IsEmailValid = blnIsItValid 
			Exit Function
		End If
		
		' split the email address in two parts: name@domain.ext 
		strArray = Split(strEmail, "@") 
	   
		' if there are more or less than two parts   
		If UBound(strArray) <> 1 Then 
			blnIsItValid = False 
			IsEmailValid = blnIsItValid 
			Exit Function 
		End If 
	   
		' check each part 
		For Each strItem In strArray 
			' no part can be void 
			If Len(strItem) <= 0 Then 
				blnIsItValid = False 
				IsEmailValid = blnIsItValid 
				Exit Function 
			End If 
	         
			' check each character of the part 
			' only following "abcdefghijklmnopqrstuvwxyz_-." 
			' characters and the ten digits are allowed 
			For i = 1 To Len(strItem) 
				c = LCase(Mid(strItem, i, 1)) 
				' if there is an illegal character in the part 
				If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then 
					blnIsItValid = False 
					IsEmailValid = blnIsItValid 
					Exit Function 
				End If 
			Next 
	    
		' the first and the last character in the part cannot be . (dot) 
			If Left(strItem, 1) = "." Or Right(strItem, 1) = "." Then 
			blnIsItValid = False 
			IsEmailValid = blnIsItValid 
			Exit Function 
			End If 
		Next 
	   
		' the second part (domain.ext) must contain a . (dot) 
		If InStr(strArray(1), ".") <= 0 Then 
			blnIsItValid = False 
			IsEmailValid = blnIsItValid 
			Exit Function 
		End If 
	   
		' check the length oh the extension   
		i = Len(strArray(1)) - InStrRev(strArray(1), ".") 
		' the length of the extension can be only 2, 3, or 4 
		' to cover the new "info" extension 
		If i <> 2 And i <> 3 And i <> 4 Then 
			blnIsItValid = False 
			IsEmailValid = blnIsItValid 
			Exit Function 
		End If 

		' after . (dot) cannot follow a . (dot) 
		If InStr(strEmail, "..") > 0 Then 
			blnIsItValid = False 
			IsEmailValid = blnIsItValid 
			Exit Function 
		End If 
	   
		' finally it's OK   
		IsEmailValid = blnIsItValid 
     
End Function 


Response.Write IsEmailValid("1234_g@walla.co.il ") 

%>
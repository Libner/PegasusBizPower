<% 
   Function hebrewDate(date)
        if date=nil then
           hebrewDate=""
        else
           hebrewDate=CStr(Day(date))&"/"&CStr(Month(date))&"/"&CStr(Year(date))
        end if
   end Function

'ORACLE Server   
Function sFix(myString)
	Dim quote, newQuote, temp
	temp = myString
	If trim(temp) <> "" Then
		quote = Chr(39)
		newQuote = Chr(39)&Chr(39)
		temp = Replace(temp, quote, newQuote)
		sFix = temp
	Else 		
		sFix = ""
	End If
End Function

Function vFix(word)
	word=replace("" & word,"<","&lt;")
	word=replace("" & word,">","&gt;")
	word=replace("" & word,"₪","&curren;")
	word=replace("" & word,chr(34),"&quot;")
	vFix=word
End Function
   
  Function altFix(word)
   dim strtemp
   if trim(word)<>nil then
     strtemp=replace(word,chr(13) & chr(10),"<br>")
     strtemp=Replace(strtemp,"'","`") 
     word=Replace(strtemp,"""","``")
     altFix=word
   else
     altFix=""
   end if  
end Function
  Function altFixRev(word)
     dim strtemp
     if trim(word)<>nil then
       strtemp=replace(word,"<br>",chr(13) & chr(10))
       strtemp=Replace(strtemp,"`","'") 
       word=Replace(strtemp,"""","``")
       altFix=word
     else
       altFix=""
     end if  
  end Function
  
   Function breaks(word)			
		if word<>nil AND word<>"" then
	     word = Replace(word, Chr(13) & Chr(10), "<br>")
                word = Replace(word, Chr(13), "<br>")
                word = Replace(word, Chr(10), "<br>")
                breaks = word
		else
		  breaks=""
		end if
		
		   End Function   
   
Function nFix(myString)
	If isNumeric(myString) Then
		nFix = cLng(myString)
	Else
		nFix = "NULL"
	End If	
End Function      

 function GiveDec(Hex)
		if Hex = "A" then
			GiveDec = 10
		elseif Hex = "B" then
			GiveDec = 11
		elseif Hex = "C" then
			GiveDec = 12
		elseif Hex = "D" then
			GiveDec = 13
		elseif Hex = "E" then
			GiveDec = 14
		elseif Hex = "F" then
			GiveDec = 15
		else
			GiveDec = Hex
		end if
end function
		
	function GiveHex(Dec)
		if Dec = 10 then
			GiveHex = "A"
		elseif Dec = 11 then
			GiveHex = "B"
		elseif Dec = 12 then
			GiveHex = "C"
		elseif Dec = 13 then
			GiveHex = "D"
		elseif Dec = 14 then
			GiveHex = "E"
		elseif Dec = 15 then
			GiveHex = "F"
		else
			GiveHex = "" & Dec
		end if	
	end function
	
	
			function calcColor(color, interval)
			If Len(color) > 0 Then
				Input_Color = Right(color,Len(color)-1)

				a = GiveDec(Mid(Input_Color,1, 1))
				b = GiveDec(Mid(Input_Color,2, 1))
				c = GiveDec(Mid(Input_Color,3, 1))
				d = GiveDec(Mid(Input_Color,4, 1))
				e = GiveDec(Mid(Input_Color,5, 1))
				f = GiveDec(Mid(Input_Color,6, 1))
				
				'Response.Write a
				'Response.End
				x = (a * 16) + b : y = (c * 16) + d : z = (e * 16) + f
				
				If x + interval > 0 Then
					x = x + interval
				Else
					x = 0	
				End If
				
				If y + interval > 0 Then
					y = y + interval
				Else
					y = 0	
				End If
				
				If z + interval > 0 Then
					z = z + interval
				Else
					z = 0	
				End If	
				
				Red = x : Green = y : Blue = z

				a = GiveHex(Fix(Red / 16))

				b = GiveHex(Red Mod 16)

				c = GiveHex(Fix(Green / 16))

				d = GiveHex(Green Mod 16)

				e = GiveHex(Fix(Blue / 16))

				f = GiveHex(Blue Mod 16)
				
				calcColor = "#" & cStr(a + b + c + d + e + f)	
			Else
				calcColor = "#000000"
			End If				
			end function	
			
	Function FormatMediumDateShort(DateValue)
		Dim strYYYY
		Dim strMM
		Dim strDD

		strYYYY = CStr(DatePart("yyyy", DateValue))
		If Len(strYYYY) > 1 Then
			strYY = Right(strYYYY,2)
		Else
			strYY = strYYYY
		End If

		strMM = CStr(DatePart("m", DateValue))
		If Len(strMM) = 1 Then strMM = "0" & strMM

		strDD = CStr(DatePart("d", DateValue))
		If Len(strDD) = 1 Then strDD = "0" & strDD

		FormatMediumDateShort = strDD & "/" & strMM & "/" & strYY

	End Function 					
	
Function GetTime(inMin)
	dim str_hour,str_min
    If cdbl(inMin) >= 60 Then
        str_hour = Right("0" & CStr(cdbl(Split(CStr(inMin / 60), ".")(0)) Mod 60),2) & ":"
    Else
        str_hour = "00:"
    End If
    str_min = Right("0" & CStr(cdbl(inMin) Mod 60),2)
    GetTime = str_hour & str_min
End Function

   Function Encode(sIn)
            Dim x, y, abfrom, abto
            Encode = "" : abfrom = ""

            For x = 0 To 25 : abfrom = abfrom & Chr(65 + x) : Next
            For x = 0 To 25 : abfrom = abfrom & Chr(97 + x) : Next
            For x = 0 To 9 : abfrom = abfrom & CStr(x) : Next

            abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
            For x = 1 To Len(sIn) : y = InStr(abfrom, Mid(sIn, x, 1))
                If y = 0 Then
                    Encode = Encode & Mid(sIn, x, 1)
                Else
                    Encode = Encode & Mid(abto, y, 1)
                End If
            Next
    End Function
    
    Function Decode(sIn)
            Dim x, y, abfrom, abto
            Decode = "" : abfrom = ""

            For x = 0 To 25 : abfrom = abfrom & Chr(65 + x) : Next
            For x = 0 To 25 : abfrom = abfrom & Chr(97 + x) : Next
            For x = 0 To 9 : abfrom = abfrom & CStr(x) : Next

            abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
            For x = 1 To Len(sIn) : y = InStr(abto, Mid(sIn, x, 1))
                If y = 0 Then
                    Decode = Decode & Mid(sIn, x, 1)
                Else
                    Decode = Decode & Mid(abfrom, y, 1)
                End If
            Next
     End Function
     
     Function DBValueToStr(Value)
		If IsNull(Value) Then
			DBValueToStr = ""
		Else
			DBValueToStr = CStr(trim(Value))
		End If
	End Function 
	  Function TourStatusForOperation(Departure_Id , Departure_DateBegin,  Departure_DateEnd) 
         '   Dim temp As String
           today  = Now()
            DateBegin = Departure_DateBegin
            DateEnd = Departure_DateEnd
            If DateDiff("d", today, DateBegin) >= 0 Then

                If DateDiff("d", today, DateBegin) = 0 Then
                    temp = "יוצא היום "
                Else
                    '  temp = "יוצא בעוד " & DateDiff(DateInterval.Day, today, DateBegin) & " ימים "
                    temp = DateDiff("d", today, DateBegin)

                End If

            ElseIf DateDiff("d", today, DateBegin) < 0 Then

                If DateDiff("d", today, DateEnd) < 0 Then
                    temp = "חזר"
                Else
                    temp = "כן"
                End If
            End If
            'temp = DateBegin & " :" & DateDiff(DateInterval.Day, today, DateBegin) & "<BR> " & DateEnd & ":" & DateDiff(DateInterval.Day, today, DateEnd)

            '  temp = Departure_DateBegin & ":" & Departure_DateEnd
            ' temp = DateDiff(DateInterval.Day, today, DateBegin)
            TourStatusForOperation= temp
        End Function
        Function dbNullFix(objVal)
            If IsNull(objVal) Then
                dbNullFix = ""
            Else
                dbNullFix = Trim(objVal)
            End If
        End Function
        Function GetVouchers_ProviderStatus(ValueStr )
        
            ss = ValueStr  
          
 If ss <> "" Then

               sql="SELECT Vouchers_Status from VouchersToSuppliers where Departure_Id =" & ss
               set rsf = con.getRecordSet(sql)
				If Not rsf.Eof Then
						 Do while Not rsf.eof 
								  If Not IsNull(rsf("Vouchers_Status")) Then
										  If Trim(rsf("Vouchers_Status")) <> "מותאם" Then
											 result = ""
											 GetVouchers_ProviderStatus= result
										  Else
											 result = "מותאם סופית"
										  End If
								  end if

						 rsf.movenext
						 Loop

				end if
               set rsf=Nothing
               GetVouchers_ProviderStatus= result
   End If
        End Function
        
		Function PermissionsByDepartmentId(UID)
					sql="SELECT PermissionsByDepartmentId  FROM Users where User_id=" & UID
					set rsf = con.getRecordSet(sql)
					If Not IsNull(rsf("PermissionsByDepartmentId")) Then
					result=rsf("PermissionsByDepartmentId")
					end if
					set rsf=Nothing
					PermissionsByDepartmentId=result
					
        End Function
        
        Function GetSelectSupplierName(ValueStr)
          
            ss = ValueStr
              If ss <> "" Then

                sql="SELECT supplier_Name from Suppliers  where  supplier_ID in(" & ss & ") order by supplier_Name"
               set rsf = con.getRecordSet(sql)
             do   While Not rsf.eof 
                    If Not IsNull(rsf("supplier_Name")) Then
                        If result <> "" Then
                            result = result & ", " & rsf("supplier_Name")
                        Else
                            result = rsf("supplier_Name")
                        End If


                    End If

                rsf.movenext
						 Loop
                'If result = "" Then
                '    result = "הכל"
                'End If
                set rsf=Nothing


                GetSelectSupplierName=result
            End If
        End Function
        
        Function GetCountGuideMessages(Departure_Id) 
                 If Departure_Id <> "" Then
				   sqlstr = "Select Count(*) as s from GuideMessages where Departure_Id=" & Departure_Id
				   set rsf = con.getRecordSet(sqlstr)
						If Not rsf.Eof Then
							 Count = rsf("s")
							  If Count > 0 Then
								  tmp = Count
							  Else
								  tmp = ""
							  End If

                         GetCountGuideMessages=tmp

                       End If
						set rsf=Nothing
						end if
						
						End Function
 Function StatusOperation(objVal) 
            If IsNull(objVal) Then
                StatusOperation = ""
            Else
                If CInt(objVal) <= 0 Then
                    StatusOperation = "כן"
                Else
                    StatusOperation = ""

                End If
            End If
            
        End Function
	%>

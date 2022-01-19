<%
 'Server.ScriptTimeout = 360
Response.Expiresabsolute = Now() - 1 
Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","private"
Response.CacheControl = "no-cache" 
set conPegasus=server.createobject("adodb.connection")
  conPegasus.Open "Provider=SQLOLEDB.1;Password=cyber;Persist Security Info=True;User ID=pegasus_user;" &_
  "Initial Catalog=pegasus_test;Data Source=192.168.0.40"  

 connString = Application("ConnectionString")
 'conPegsus=	Application("ConnectionStringPegasus") 
 Session.LCID = 2057
 Response.CharSet = "windows-1255"
 set con = new BizCon

strLocal = "http://" & Request.ServerVariables("SERVER_NAME")
If Len(Request.ServerVariables("SERVER_PORT")) > 0 Then
	strLocal = strLocal & ":" & Request.ServerVariables("SERVER_PORT")
End If
strLocal = strLocal & Application("VirDir") & "/"

'added by Mila 16/07/2017 - DB names for using in joined queries between two databases
'!!!!Change in PRODUCTION!!!!!!!!!!!!!!!
pegasusDBName=Application("pegasusDBName")
bizpowerDBName=Application("bizpowerDBName")

Class BizCon

Dim conn

Function GetRecordSet(pSqlString)  
'response.write   pSqlString
'response.end
    Set l_rsData = Server.CreateObject("ADODB.Recordset")
    l_rsData.CursorLocation = 3
'On Error Resume Next
 'response.Write pSqlString
    l_rsData.Open CStr(pSqlString), conn
    
    If Err.number = 0 Then
		Set GetRecordSet = l_rsData
    Else
		'Response.Write pSqlString
		Set GetRecordSet = Nothing
		Err.Clear 
    End If
    
    Set l_rsData = Nothing
End Function

Function ExecuteQuery(pSqlString)
' On Error Resume Next
 'Response.Write pSqlString
  '  response.end
    conn.Execute CStr(pSqlString)
    
    If conn.Errors.Count > 0 Then
		'Response.Write pSqlString
        ExecuteQuery = False
        'Response.End 
    Else
        ExecuteQuery = True
    End If

End Function

Private Sub Class_Initialize()
   
    Set conn = Server.CreateObject("ADODB.Connection")   
    conn.CommandTimeout = 120  ' Default is 30
    conn.ConnectionTimeout =120 '  'default=15sec
    conn.Open connString
    
End Sub

Private Sub Class_Terminate()
    conn.Close
    Set conn = Nothing
End Sub

Function GetFormField(pField_Id, pField_Type, pField_Size, pField_Align, pField_Scale, pPathCalImage, pLanguage, pField_Value, pField_Title)
Dim tmpStr, selValue, i, strValue, isChecked, isSelected
Set rsSelect = Server.CreateObject("ADODB.Recordset")
Dim tmparr, strImportance, impH, impM, impL, impNA

   strValue = pField_Value
    If IsNumeric(pField_Size) Then
        Field_Size = CInt(pField_Size)
        If Field_Size > 50 Then
            Field_Size = 50
        End If
    End If
    Select Case pField_Type
    Case "1" ' one string text
        GetFormField = "<INPUT type='text' dir=" & pField_Align & " maxlength=" & pField_Size & " size=" & Field_Size & " id=field" & pField_Id & " name=field" & pField_Id & " value=""" & Replace(strValue, Chr(34), "''") & """>"
    Case "2" 'textarea
        GetFormField = "<textarea dir=" & pField_Align & " class='txt' rows=5 cols=30 style='width: 220px' id=field" & pField_Id & " name=field" & pField_Id & ">" & strValue & "</textarea>"
    Case "3" 'select
        tmpStr = "<select dir=" & pField_Align & " class='norm' id=field" & pField_Id & " name=field" & pField_Id & ">" & vbCrLf
        tmpStr = tmpStr & "<option value=''>"
        If pLanguage = "eng" Then
            tmpStr = tmpStr & " Choose "
        Else
            tmpStr = tmpStr & " בחר תשובה "
        End If
        tmpStr = tmpStr & "</option>" & vbCrLf
        if Field_id=40663 then 'yadim        
        Set rsSelect = conn.Execute("Select FIELD_VALUE from Form_select where Field_id=" & pField_Id & " Order by FIELD_VALUE")
        else
        Set rsSelect = conn.Execute("Select FIELD_VALUE from Form_select where Field_id=" & pField_Id & " Order by Id")
        end if
        If not rsSelect.Eof Then
			arr_values = rsSelect.getRows()
        End If
        Set rsSelect = Nothing
        
        If isArray(arr_values) Then
			For vv=0 To Ubound(arr_values,2)
                selValue = trim(arr_values(0,vv))
                selValue = Replace(selValue, Chr(34), "''")
                If selValue = strValue Then
                    isSelected = " selected"
                Else
                    isSelected = ""
                End If
                tmpStr = tmpStr & "<option value=""" & selValue & """ " & isSelected & ">" & selValue & "</option>" & vbCrLf
			Next
        End If
        tmpStr = tmpStr & "</select>"
        GetFormField = tmpStr
    Case "4" 'file attachment
        If strValue <> "" Then
        strValue = Replace(strValue, "'", "&#39;")
         GetFormField = "<a href='/download/forms_attachments/" & strValue & "' target='_blank' class='file_link' >קובץ מצורף</a>" & _
         "<br><INPUT type='file' id=field" & pField_Id & " name=field" & pField_Id & " dir='ltr' size='20' >"
        Else
        GetFormField = "<INPUT type='file' id=field" & pField_Id & " name=field" & pField_Id & " dir='ltr' size='20' >"
        End If
      
    Case "5" 'check
        If strValue = "yes" Or strValue = "כן" Or strValue = "on" Then
            isChecked = " checked"
        Else
            isChecked = ""
        End If
        GetFormField = "<INPUT type=checkbox id=field" & pField_Id & " name=field" & pField_Id & " " & isChecked & ">"
    Case "6" 'date
        If strValue <> "" Then
            isChecked = strValue
        Else
            isChecked = vbNullString
        End If
        If pLanguage = "eng" Then
            GetFormField = "<INPUT type=text dir=ltr class=passw maxlength=10 id=field" & pField_Id & " name=field" & pField_Id & " size=10 value='" & isChecked & "' onclick='cal.select(this,this.id,""dd/MM/yyyy""); return false;' readonly>&nbsp;<Img class='hand' src='" & pPathCalImage & "images/calend.gif' id=imgcal name=imgcal border=0 onclick='cal.select(window.document.getElementById(""field" & pField_Id & """), ""field" & pField_Id & """, ""dd/MM/yyyy""); return false;'>"
        Else
            GetFormField = "<Img class='hand' src='../../images/calend.gif' id=imgcal name=imgcal border=0 onclick='cal.select(window.document.getElementById(""field" & pField_Id & """), ""field" & pField_Id & """, ""dd/MM/yyyy""); return false;'>&nbsp;<INPUT type=text dir=ltr class=passw maxlength=10 id=field" & pField_Id & " name=field" & pField_Id & " size=10 value='" & isChecked & "' onclick='cal.select(this,this.id,""dd/MM/yyyy""); return false;' readonly>"
        End If
    Case "7" 'number
        GetFormField = "<INPUT type=text dir=ltr class='texts' maxlength=" & pField_Size & " id=field" & pField_Id & " name=field" & pField_Id & " value=""" & Replace(strValue, Chr(34), "''") & """>"
    Case "8" ' radio - select
      If pLanguage = "eng" Then
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=ltr>" & vbCrLf
        Else
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=rtl>" & vbCrLf
        End If
        
        Set rsSelect = conn.Execute("Select Field_Value from Form_select where Field_id=" & pField_Id & " Order by Id")
        If not rsSelect.Eof Then
			arr_values = rsSelect.getRows()
        End If
        Set rsSelect = Nothing
        
        If isArray(arr_values) Then
			For vv=0 To Ubound(arr_values,2)
                selValue = trim(arr_values(0,vv))
                selValue = Replace(selValue, Chr(34), "''")
                If selValue = strValue Then
                    isChecked = " checked"
                Else
                    isChecked = ""
                End If
                tmpStr = tmpStr & "<TR>" & vbCrLf
                If pLanguage = "eng" Then
                    tmpStr = tmpStr & "<TD><INPUT type=radio id='field" & pField_Id & "_" & vv & "' name='field" & pField_Id & "' value=""" & selValue & """ " & isChecked & "></TD><Td class='form' valign=top>" & selValue & "</Td><td width=5>&nbsp;</td>" & vbCrLf
                Else
                    tmpStr = tmpStr & "<td width=5>&nbsp;</td><TD><INPUT type=radio id='field" & pField_Id & "_" & vv & "' name='field" & pField_Id & "' value=""" & selValue & """ " & isChecked & "></TD><Td class='form' dir=rtl valign=top>" & selValue & "</td>" & vbCrLf
                End If
                tmpStr = tmpStr & "</TR>" & vbCrLf
			Next
		End If       
        tmpStr = tmpStr & "</TABLE>"
        GetFormField = tmpStr
    Case "9" 'scale
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1>" & vbCrLf
        For i = 1 To CInt(pField_Scale)
            If CStr(i) = strValue Then
                isChecked = " checked"
            Else
                isChecked = ""
            End If
            tmpStr = tmpStr & "<TR>" & vbCrLf
            If pLanguage = "eng" Then
                tmpStr = tmpStr & "<TD><INPUT type=radio id='field" & pField_Id & "_" & i & "' name='field" & pField_Id & "' value=" & i & " " & isChecked & "></TD><Td class='form' valign=top>" & i & "</Td><td width=5>&nbsp;</td>" & vbCrLf
            Else
                tmpStr = tmpStr & "<td width=5>&nbsp;</td><Td class='form' valign='middle'>" & i & "</Td><TD><INPUT type=radio id='field" & pField_Id & "_" & i & "' name='field" & pField_Id & "' value=" & i & " " & isChecked & "></TD>" & vbCrLf
            End If
            tmpStr = tmpStr & "</TR>" & vbCrLf
        Next
        'If strValue = "" Then
        '    isChecked = " checked"
        'Else
        '    isChecked = ""
        'End If
        'tmpStr = tmpStr & "<TR>" & vbCrLf
        'If pLanguage = "eng" Then
        '    tmpStr = tmpStr & "<TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value="""" " & isChecked & "></TD><Td class='form' valign=top>Not relevant</Td><td width=5>&nbsp;</td>" & vbCrLf
        'Else
        '    tmpStr = tmpStr & "<td width=5>&nbsp;</td><Td class='form' valign=top>לא רלונטי</Td><TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value="""" " & isChecked & "></TD>" & vbCrLf
        'End If
        'tmpStr = tmpStr & "</TR>" & vbCrLf
        tmpStr = tmpStr & "</TABLE>"
        GetFormField = tmpStr
    'Case "10" - subject - can't use here
    Case "11" ' radio - select - vertical
        If pLanguage = "eng" Then
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=ltr>" & vbCrLf
        Else
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=rtl>" & vbCrLf
        End If
        tmpStr = tmpStr & "<TR>" & vbCrLf
        
        Set rsSelect = conn.Execute("Select Field_Value from Form_select where Field_id=" & pField_Id & " Order by Id")
        If not rsSelect.Eof Then
			arr_values = rsSelect.getRows()
        End If
        Set rsSelect = Nothing
        
        If isArray(arr_values) Then
			For vv=0 To Ubound(arr_values,2)
                selValue = trim(arr_values(0,vv))
                selValue = Replace(selValue, Chr(34), "''")
                If selValue = strValue Then
                    isChecked = " checked"
                Else
                    isChecked = ""
                End If
                If pLanguage = "eng" Then
                    tmpStr = tmpStr & "<TD><INPUT type=radio id='field" & pField_Id & "_" & vv & "' name='field" & pField_Id & "' value=""" & selValue & """ " & isChecked & "></TD><Td class='form' valign=top>" & selValue & "</Td><td width=5>&nbsp;</td>" & vbCrLf
                Else
                    tmpStr = tmpStr & "<td width=5>&nbsp;</td><TD valign='middle'><INPUT type=radio id='field" & pField_Id & "_" & vv & "' name='field" & pField_Id & "' value=""" & selValue & """ " & isChecked & "></TD><Td class='form' dir=rtl valign='middle'>" & selValue & "</Td>" & vbCrLf
                End If
			Next
		End If	
        tmpStr = tmpStr & "</TR>" & vbCrLf
        tmpStr = tmpStr & "</TABLE>"
        GetFormField = tmpStr
      Case "12" 'scale - vertical
      If pLanguage = "eng" Then
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=ltr>" & vbCrLf
        Else
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=rtl>" & vbCrLf
        End If
        tmpStr = tmpStr & "<TR>" & vbCrLf
        For i = 1 To CInt(pField_Scale)
            If CStr(i) = strValue Then
                isChecked = " checked"
            Else
                isChecked = ""
            End If
            If pLanguage = "eng" Then
                tmpStr = tmpStr & "<TD><INPUT type=radio id='field" & pField_Id & "_" & i & "' name='field" & pField_Id & "' value=" & i & " " & isChecked & "></TD><Td class='form' valign=top>" & i & "</Td><td width=5>&nbsp;</td>" & vbCrLf
            Else
                tmpStr = tmpStr & "<td width=5>&nbsp;</td><TD valign='top'><INPUT type=radio id='field" & pField_Id & "_" & i & "' name='field" & pField_Id & "' value=" & i & " " & isChecked & "></TD><Td class='form' valign=top>" & i & "</Td>" & vbCrLf
            End If
        Next
        'If strValue = "" Then
        '    isChecked = " checked"
        'Else
        '    isChecked = ""
        'End If
        'If pLanguage = "eng" Then
        '    tmpStr = tmpStr & "<TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value="""" " & isChecked & "></TD><Td class='form' valign=top>Not relevant</Td><td width=10>&nbsp;</td>" & vbCrLf
        'Else
        '    tmpStr = tmpStr & "<td width=10>&nbsp;</td><TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value="""" " & isChecked & "></TD><Td class='form' valign=top>לא רלונטי</Td>" & vbCrLf
        'End If
        
        tmpStr = tmpStr & "</TR>" & vbCrLf
        
     
        tmpStr = tmpStr & "</TABLE>"
        GetFormField = tmpStr
     End Select
     
 
End Function
Function GetFormFieldNoEdit(pField_Id, pField_Type, pField_Size, pField_Align, pField_Scale, pPathCalImage, pLanguage, pField_Value, pField_Title)
Dim tmpStr, selValue, i, strValue, isChecked, isSelected
Set rsSelect = Server.CreateObject("ADODB.Recordset")
Dim tmparr, strImportance, impH, impM, impL, impNA

   strValue = pField_Value
    If IsNumeric(pField_Size) Then
        Field_Size = CInt(pField_Size)
        If Field_Size > 50 Then
            Field_Size = 50
        End If
    End If
    Select Case pField_Type
    Case "1" ' one string text
        GetFormFieldNoEdit = "<INPUT type='hidden' dir=" & pField_Align & " maxlength=" & pField_Size & " size=" & Field_Size & " id=field" & pField_Id & " name=field" & pField_Id & " value=""" & Replace(strValue, Chr(34), "''") & """>"& "<div id=fieldNoEdit" & pField_Id & ">" & strValue & "</div>"
    Case "2" 'textarea
        GetFormFieldNoEdit = "<textarea dir=" & pField_Align & " class='txt' rows=5 cols=30 style='width: 220px' id=field" & pField_Id & " name=field" & pField_Id & ">" & strValue & "</textarea>"
    Case "3" 'select
        tmpStr = "<select dir=" & pField_Align & " class='norm' id=field" & pField_Id & " name=field" & pField_Id & ">" & vbCrLf
        tmpStr = tmpStr & "<option value=''>"
        If pLanguage = "eng" Then
            tmpStr = tmpStr & " Choose "
        Else
            tmpStr = tmpStr & " בחר תשובה "
        End If
        tmpStr = tmpStr & "</option>" & vbCrLf
        Set rsSelect = conn.Execute("Select FIELD_VALUE from Form_select where Field_id=" & pField_Id & " Order by Id")
        If not rsSelect.Eof Then
			arr_values = rsSelect.getRows()
        End If
        Set rsSelect = Nothing
        
        If isArray(arr_values) Then
			For vv=0 To Ubound(arr_values,2)
                selValue = trim(arr_values(0,vv))
                selValue = Replace(selValue, Chr(34), "''")
                If selValue = strValue Then
                    isSelected = " selected"
                Else
                    isSelected = ""
                End If
                tmpStr = tmpStr & "<option value=""" & selValue & """ " & isSelected & ">" & selValue & "</option>" & vbCrLf
			Next
        End If
        tmpStr = tmpStr & "</select>"
        GetFormField = tmpStr
    Case "4" 'file attachment
        If strValue <> "" Then
        strValue = Replace(strValue, "'", "&#39;")
         GetFormField = "<a href='/download/forms_attachments/" & strValue & "' target='_blank' class='file_link' >קובץ מצורף</a>" & _
         "<br><INPUT type='file' id=field" & pField_Id & " name=field" & pField_Id & " dir='ltr' size='20' >"
        Else
        GetFormField = "<INPUT type='file' id=field" & pField_Id & " name=field" & pField_Id & " dir='ltr' size='20' >"
        End If
      
    Case "5" 'check
        If strValue = "yes" Or strValue = "כן" Or strValue = "on" Then
            isChecked = " checked"
        Else
            isChecked = ""
        End If
        GetFormField = "<INPUT type=checkbox id=field" & pField_Id & " name=field" & pField_Id & " " & isChecked & ">"
    Case "6" 'date
        If strValue <> "" Then
            isChecked = strValue
        Else
            isChecked = vbNullString
        End If
        If pLanguage = "eng" Then
            GetFormField = "<INPUT type=text dir=ltr class=passw maxlength=10 id=field" & pField_Id & " name=field" & pField_Id & " size=10 value='" & isChecked & "' onclick='cal.select(this,this.id,""dd/MM/yyyy""); return false;' readonly>&nbsp;<Img class='hand' src='" & pPathCalImage & "images/calend.gif' id=imgcal name=imgcal border=0 onclick='cal.select(window.document.getElementById(""field" & pField_Id & """), ""field" & pField_Id & """, ""dd/MM/yyyy""); return false;'>"
        Else
            GetFormField = "<Img class='hand' src='../../images/calend.gif' id=imgcal name=imgcal border=0 onclick='cal.select(window.document.getElementById(""field" & pField_Id & """), ""field" & pField_Id & """, ""dd/MM/yyyy""); return false;'>&nbsp;<INPUT type=text dir=ltr class=passw maxlength=10 id=field" & pField_Id & " name=field" & pField_Id & " size=10 value='" & isChecked & "' onclick='cal.select(this,this.id,""dd/MM/yyyy""); return false;' readonly>"
        End If
    Case "7" 'number
        GetFormField = "<INPUT type=text dir=ltr class='texts' maxlength=" & pField_Size & " id=field" & pField_Id & " name=field" & pField_Id & " value=""" & Replace(strValue, Chr(34), "''") & """>"
    Case "8" ' radio - select
      If pLanguage = "eng" Then
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=ltr>" & vbCrLf
        Else
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=rtl>" & vbCrLf
        End If
        
        Set rsSelect = conn.Execute("Select Field_Value from Form_select where Field_id=" & pField_Id & " Order by Id")
        If not rsSelect.Eof Then
			arr_values = rsSelect.getRows()
        End If
        Set rsSelect = Nothing
        
        If isArray(arr_values) Then
			For vv=0 To Ubound(arr_values,2)
                selValue = trim(arr_values(0,vv))
                selValue = Replace(selValue, Chr(34), "''")
                If selValue = strValue Then
                    isChecked = " checked"
                Else
                    isChecked = ""
                End If
                tmpStr = tmpStr & "<TR>" & vbCrLf
                If pLanguage = "eng" Then
                    tmpStr = tmpStr & "<TD><INPUT type=radio id='field" & pField_Id & "_" & vv & "' name='field" & pField_Id & "' value=""" & selValue & """ " & isChecked & "></TD><Td class='form' valign=top>" & selValue & "</Td><td width=5>&nbsp;</td>" & vbCrLf
                Else
                    tmpStr = tmpStr & "<td width=5>&nbsp;</td><TD><INPUT type=radio id='field" & pField_Id & "_" & vv & "' name='field" & pField_Id & "' value=""" & selValue & """ " & isChecked & "></TD><Td class='form' dir=rtl valign=top>" & selValue & "</td>" & vbCrLf
                End If
                tmpStr = tmpStr & "</TR>" & vbCrLf
			Next
		End If       
        tmpStr = tmpStr & "</TABLE>"
        GetFormField = tmpStr
    Case "9" 'scale
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1>" & vbCrLf
        For i = 1 To CInt(pField_Scale)
            If CStr(i) = strValue Then
                isChecked = " checked"
            Else
                isChecked = ""
            End If
            tmpStr = tmpStr & "<TR>" & vbCrLf
            If pLanguage = "eng" Then
                tmpStr = tmpStr & "<TD><INPUT type=radio id='field" & pField_Id & "_" & i & "' name='field" & pField_Id & "' value=" & i & " " & isChecked & "></TD><Td class='form' valign=top>" & i & "</Td><td width=5>&nbsp;</td>" & vbCrLf
            Else
                tmpStr = tmpStr & "<td width=5>&nbsp;</td><Td class='form' valign='middle'>" & i & "</Td><TD><INPUT type=radio id='field" & pField_Id & "_" & i & "' name='field" & pField_Id & "' value=" & i & " " & isChecked & "></TD>" & vbCrLf
            End If
            tmpStr = tmpStr & "</TR>" & vbCrLf
        Next
         tmpStr = tmpStr & "</TABLE>"
        GetFormField = tmpStr
    'Case "10" - subject - can't use here
    Case "11" ' radio - select - vertical
        If pLanguage = "eng" Then
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=ltr>" & vbCrLf
        Else
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=rtl>" & vbCrLf
        End If
        tmpStr = tmpStr & "<TR>" & vbCrLf
        
        Set rsSelect = conn.Execute("Select Field_Value from Form_select where Field_id=" & pField_Id & " Order by Id")
        If not rsSelect.Eof Then
			arr_values = rsSelect.getRows()
        End If
        Set rsSelect = Nothing
        
        If isArray(arr_values) Then
			For vv=0 To Ubound(arr_values,2)
                selValue = trim(arr_values(0,vv))
                selValue = Replace(selValue, Chr(34), "''")
                If selValue = strValue Then
                    isChecked = " checked"
                Else
                    isChecked = ""
                End If
                If pLanguage = "eng" Then
                    tmpStr = tmpStr & "<TD><INPUT type=radio id='field" & pField_Id & "_" & vv & "' name='field" & pField_Id & "' value=""" & selValue & """ " & isChecked & "></TD><Td class='form' valign=top>" & selValue & "</Td><td width=5>&nbsp;</td>" & vbCrLf
                Else
                    tmpStr = tmpStr & "<td width=5>&nbsp;</td><TD valign='middle'><INPUT type=radio id='field" & pField_Id & "_" & vv & "' name='field" & pField_Id & "' value=""" & selValue & """ " & isChecked & "></TD><Td class='form' dir=rtl valign='middle'>" & selValue & "</Td>" & vbCrLf
                End If
			Next
		End If	
        tmpStr = tmpStr & "</TR>" & vbCrLf
        tmpStr = tmpStr & "</TABLE>"
        GetFormField = tmpStr
      Case "12" 'scale - vertical
      If pLanguage = "eng" Then
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=ltr>" & vbCrLf
        Else
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=rtl>" & vbCrLf
        End If
        tmpStr = tmpStr & "<TR>" & vbCrLf
        For i = 1 To CInt(pField_Scale)
            If CStr(i) = strValue Then
                isChecked = " checked"
            Else
                isChecked = ""
            End If
            If pLanguage = "eng" Then
                tmpStr = tmpStr & "<TD><INPUT type=radio id='field" & pField_Id & "_" & i & "' name='field" & pField_Id & "' value=" & i & " " & isChecked & "></TD><Td class='form' valign=top>" & i & "</Td><td width=5>&nbsp;</td>" & vbCrLf
            Else
                tmpStr = tmpStr & "<td width=5>&nbsp;</td><TD valign='top'><INPUT type=radio id='field" & pField_Id & "_" & i & "' name='field" & pField_Id & "' value=" & i & " " & isChecked & "></TD><Td class='form' valign=top>" & i & "</Td>" & vbCrLf
            End If
        Next
    
        
        tmpStr = tmpStr & "</TR>" & vbCrLf
        
     
        tmpStr = tmpStr & "</TABLE>"
        GetFormFieldNoEdit = tmpStr
     End Select
     
 
End Function


Function GetFieldGraph(pProduct_Id, pField_Id, pField_Title, pField_Type, pField_Scale, pLanguage, pType_ID)
Dim tmpStr, selValue, colind, selAll, wd, i
Dim td_align
Dim td_width
Dim ScQuant
Set rsSelect = Server.CreateObject("ADODB.Recordset")
Dim col(9)
col(0) = "#65A825"
col(1) = "#19B4D0"
col(2) = "#E87117"
col(3) = "#6E15C3"
col(4) = "#44B4AD"
col(5) = "#011D66"
col(6) = "#F84637"
col(7) = "#1DF3FD"
col(8) = "#EBB20D"
colind = 0
    If pLanguage = "eng" Then
        td_align = "left"
    Else
        td_align = "right"
        string_sort = "DESC"
    End If
    Select Case pField_Type
    Case "1" ' one string text
        GetFieldGraph = ""
    Case "2" 'textarea
        GetFieldGraph = ""
    Case "3", "8", "11" 'select
        Set rsSelect = GetRecordSet("Select FIELD_SELECT, FIELD_QUANTITY from FIELD_STATISTICS where Product_Id = " & pProduct_Id & " and  Field_id=" & pField_Id & " and Type_ID = " & pType_ID & " Order by Field_Order " & string_sort)
        If Not rsSelect.Eof Then
			arr_graph = rsSelect.getRows()
        End If
        Set rsSelect = Nothing        
        
        If isArray(arr_graph) Then
        
			recCount = Ubound(arr_graph,2)
			
			selAll = 0
			tmpStr = "<tr><td bgcolor=#DADADA class='form' align=" & td_align & " nowrap valign=top colspan=" &_
			recCount & "><b>&nbsp;" & pField_Title & "&nbsp;</td></tr>" & vbCrLf
        
			If recCount <> 0 Then
				td_width = 100 / recCount
			End If        
			
            tmpStr = tmpStr & "<tr>" & vbCrLf
            For gg=0 To recCount
                selAll = selAll + CInt(arr_graph(1,gg))            
            Next
            
            For gg=0 To recCount
                If selAll <> 0 Then
                    wd = CInt((CInt(arr_graph(1,gg)) / selAll) * 100)
                    If arr_graph(1,gg) <> 0 Then
                        str_FIELD_QUANTITY = " &nbsp;<span style=color:blue;font-size:10pt>(" & arr_graph(1,gg) & ")</span>"
                    Else
                        str_FIELD_QUANTITY = ""
                    End If
                Else
                    wd = 0
                End If
                tmpStr = tmpStr & "<td align=center valign=bottom width='" & td_width & "%' height=120>" & vbCrLf
                tmpStr = tmpStr & "<TABLE width='100%' align=center BORDER=0 CELLSPACING=0 CELLPADDING=0>" & vbCrLf
                tmpStr = tmpStr & "<tr><td align=center><font size=2><font color=red>" & wd & "</font>%</font>" & str_FIELD_QUANTITY & "</td></tr>" & vbCrLf
                tmpStr = tmpStr & "<tr><td align=center valign=bottom><table border=0 width='70%' bgcolor=" & col(colind Mod 9) & " cellspacing=0 cellpadding=0>" & vbCrLf
                tmpStr = tmpStr & "<tr><td width=20 bgcolor=" & col(colind Mod 9) & " height=" & wd & "></td></tr>" & vbCrLf
                tmpStr = tmpStr & "</table></td></tr></table></td>" & vbCrLf
           
				colind = colind + 1
            Next
            
            tmpStr = tmpStr & "</tr><tr>" & vbCrLf
            
            For gg=0 To recCount
                tmpStr = tmpStr & "<td align=center bgcolor=#DADADA>&nbsp;<font size=2><b>" & arr_graph(0,gg) & "</font>&nbsp;</td>"
            Next
            tmpStr = tmpStr & "</tr>"
            arr_graph = Nothing
        End If       
       
        GetFieldGraph = tmpStr
    Case "9", "12" 'radio
        Set rsSelect = GetRecordSet("SELECT * FROM QUESTION_VALUES_VIEW WHERE Product_Id = " & pProduct_Id & " and Field_Id=" & pField_Id & " and Type_ID = " & pType_ID & " and FIELD_VALUE is not null")
        selAll = rsSelect.RecordCount
            
        tmpStr = "<tr><td bgcolor=#DADADA class='form'  align=" & td_align & " nowrap valign=top colspan=" & pField_Scale & "><b>&nbsp;" & pField_Title & "&nbsp;</td></tr>" & vbCrLf
        tmpStr = tmpStr & "<tr>" & vbCrLf
        
        If pField_Scale <> 0 Then
            td_width = 100 / pField_Scale
        End If
        
        If pLanguage = "heb" Then
            fori = CInt(pField_Scale)
            fotto = 1
            forstep = -1
        Else
            fori = 1
            fotto = CInt(pField_Scale)
            forstep = 1
        End If
        
        For i = fori To fotto Step forstep
            rsSelect.Filter = "FIELD_VALUE='" & i & "'"
            ScQuant = rsSelect.RecordCount
            If ScQuant <> 0 Then
                wd = CInt((ScQuant / selAll) * 100)
                str_FIELD_QUANTITY = "&nbsp;<span style=color:blue;font-size:10pt>(" & ScQuant & ")</span>"
            Else
                wd = 0
                str_FIELD_QUANTITY = ""
            End If
            tmpStr = tmpStr & "<td align=center valign=bottom width='" & td_width & "%' height=120>" & vbCrLf
            tmpStr = tmpStr & "<TABLE width='100%' align=center BORDER=0 CELLSPACING=0 CELLPADDING=0>" & vbCrLf
            tmpStr = tmpStr & "<tr><td align=center><font size=2><font color=red>" & wd & "</font>%</font>" & str_FIELD_QUANTITY & "</td></tr>" & vbCrLf
            tmpStr = tmpStr & "<tr><td align=center valign=bottom><table border=0 width='70%' bgcolor=" & col(colind Mod 9) & " cellspacing=0 cellpadding=0>" & vbCrLf
            tmpStr = tmpStr & "<tr><td width=20 bgcolor=" & col(colind Mod 9) & " height=" & wd & "></td></tr>" & vbCrLf
            tmpStr = tmpStr & "</table></td></tr></table></td>" & vbCrLf
        colind = colind + 1
        Next
        Set rsSelect = Nothing
        tmpStr = tmpStr & "</tr><tr>" & vbCrLf
        For i = fori To fotto Step forstep
            tmpStr = tmpStr & "<td align=center bgcolor=#DADADA>&nbsp;<font size=2><b>" & i & "</font>&nbsp;</td>"
        Next
        tmpStr = tmpStr & "</tr>"
        
        GetFieldGraph = tmpStr
    Case "5" 'check
        Set rsSelect = GetRecordSet("SELECT * FROM QUESTION_VALUES_VIEW WHERE Product_Id = " & pProduct_Id & " and Field_Id=" & pField_Id & " and Type_ID = " & pType_ID)
        selAll = rsSelect.RecordCount
        tmpStr = "<tr><td bgcolor=#DADADA class='form'  align=" & td_align & " nowrap valign=top colspan=2><b>&nbsp;" & pField_Title & "&nbsp;</td></tr>" & vbCrLf
        tmpStr = tmpStr & "<tr>" & vbCrLf
        rsSelect.Filter = "FIELD_VALUE='on'"
        ScQuant = rsSelect.RecordCount
        Set rsSelect = Nothing
        If ScQuant <> 0 Then
            wd = CInt((ScQuant / selAll) * 100)
            str_FIELD_QUANTITY = "&nbsp;<span style=color:blue;font-size:10pt>(" & ScQuant & ")</span>"
        Else
            wd = 0
            str_FIELD_QUANTITY = ""
        End If
        tmpStr = tmpStr & "<td align=center valign=bottom width='50%' height=120>" & vbCrLf
        tmpStr = tmpStr & "<TABLE width='100%' align=center BORDER=0 CELLSPACING=0 CELLPADDING=0>" & vbCrLf
        tmpStr = tmpStr & "<tr><td align=center><font size=2><font color=red>" & wd & "</font>%</font>" & str_FIELD_QUANTITY & "</td></tr>" & vbCrLf
        tmpStr = tmpStr & "<tr><td align=center valign=bottom><table border=0 width='70%' bgcolor=" & col(colind Mod 9) & " cellspacing=0 cellpadding=0>" & vbCrLf
        tmpStr = tmpStr & "<tr><td width=20 bgcolor=" & col(colind Mod 9) & " height=" & wd & "></td></tr>" & vbCrLf
        tmpStr = tmpStr & "</table></td></tr></table></td>" & vbCrLf
        
        ScQuant = selAll - ScQuant
        If ScQuant <> 0 Then
            wd = CInt((ScQuant / selAll) * 100)
            str_FIELD_QUANTITY = "&nbsp;<span style=color:blue;font-size:10pt>(" & ScQuant & ")</span>"
        Else
            wd = 0
            str_FIELD_QUANTITY = ""
        End If
        colind = colind + 1
            tmpStr = tmpStr & "<td align=center valign=bottom height=120>" & vbCrLf
        tmpStr = tmpStr & "<TABLE width='100%' align=center BORDER=0 CELLSPACING=0 CELLPADDING=0>" & vbCrLf
        tmpStr = tmpStr & "<tr><td align=center><font size=2><font color=red>" & wd & "</font>%</font>" & str_FIELD_QUANTITY & "</td></tr>" & vbCrLf
        tmpStr = tmpStr & "<tr><td align=center valign=bottom><table border=0 width='70%' bgcolor=" & col(colind Mod 9) & " cellspacing=0 cellpadding=0>" & vbCrLf
        tmpStr = tmpStr & "<tr><td width=20 bgcolor=" & col(colind Mod 9) & " height=" & wd & "></td></tr>" & vbCrLf
        tmpStr = tmpStr & "</table></td></tr></table></td>" & vbCrLf
        tmpStr = tmpStr & "</tr><tr><td align=center nowrap bgcolor=#DADADA>&nbsp;<font size=2><b>כן</font>&nbsp;</td><td align=center nowrap bgcolor=#DADADA>&nbsp;<font size=2><b>לא</font>&nbsp;</td></tr>"
         
        GetFieldGraph = tmpStr
    Case "6" 'date
        GetFieldGraph = ""
    Case "7" 'number
        GetFieldGraph = ""        
    End Select
    

End Function

End Class%>

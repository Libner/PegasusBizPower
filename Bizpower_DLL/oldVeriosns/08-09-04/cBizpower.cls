VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBizpower"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim conn As ADODB.Connection

Public Function GetRecordSet(pSqlString As Variant) As Recordset
    
    Dim l_rsData As ADODB.Recordset
    Dim l_ConnectStr  As String
    
    Set l_rsData = New ADODB.Recordset
    l_rsData.CursorLocation = adUseClient
    l_rsData.Open CStr(pSqlString), conn
    Set GetRecordSet = l_rsData
    
    Set l_rsData = Nothing
End Function

Public Function ExecuteQuery(pSqlString As Variant) As Boolean
    
    conn.Execute CStr(pSqlString)
    
    If conn.Errors.Count > 0 Then
        ExecuteQuery = False
    Else
        ExecuteQuery = True
    End If

End Function

Private Sub Class_Initialize()
    Dim l_ConnectStr  As String
    Set conn = New ADODB.Connection
    l_ConnectStr = "FILE NAME=" & App.Path & "\netReply.udl"
    conn.CommandTimeout = 0
    conn.ConnectionTimeout = 0
    conn.Open l_ConnectStr
End Sub

Private Sub Class_Terminate()
    Set conn = Nothing
End Sub

Public Function GetFormField(pField_Id, pField_Type, pField_Size, pField_Align, pField_Scale, pPathCalImage, pLanguage, Optional pField_Value As String, Optional pField_Title) As String
Dim tmpStr As String, selValue As String, i As Integer, strValue As String, isChecked As String, isSelected As String
Dim rsSelect As Recordset
Dim tmparr, strImportance As String, impH As String, impM As String, impL As String, impNA As String

   strValue = pField_Value
    If IsNumeric(pField_Size) Then
        Field_Size = CInt(pField_Size)
        If Field_Size > 50 Then
            Field_Size = 50
        End If
    End If
    Select Case pField_Type
    Case "1" ' one string text
        GetFormField = "<INPUT type='text' dir=" & pField_Align & " maxlength=" & pField_Size & " size=" & Field_Size & " id=field" & pField_Id & " name=field" & pField_Id & " value='" & strValue & "'>"
    Case "2" 'textarea
        GetFormField = "<textarea dir=" & pField_Align & " class='txt' rows=5 cols=40 id=field" & pField_Id & " name=field" & pField_Id & ">" & strValue & "</textarea>"
    Case "3" 'select
        tmpStr = "<select dir=" & pField_Align & " class='norm' id=field" & pField_Id & " name=field" & pField_Id & ">" & vbCrLf
        tmpStr = tmpStr & "<option value=''>"
        If pLanguage = "eng" Then
            tmpStr = tmpStr & " choose answer "
        Else
            tmpStr = tmpStr & " בחר תשובה "
        End If
        tmpStr = tmpStr & "</option>" & vbCrLf
        Set rsSelect = conn.Execute("Select * from Form_select where Field_id=" & pField_Id & " Order by Id")
            Do While Not rsSelect.EOF
                selValue = rsSelect("Field_Value")
                If selValue = strValue Then
                    isSelected = " selected"
                Else
                    isSelected = ""
                End If
                tmpStr = tmpStr & "<option value='" & selValue & "' " & isSelected & ">" & selValue & "</option>" & vbCrLf
            rsSelect.MoveNext
            Loop
        Set rsSelect = Nothing
        tmpStr = tmpStr & "</select>"
        GetFormField = tmpStr
    Case "4" 'scale with degree of importance
        If strValue <> "" Then
            tmparr = Split(strValue)
            strValue = tmparr(0)
            If UBound(tmparr) > 0 Then
                strImportance = tmparr(1)
            Else
                If pLanguage = "eng" Then
                    strImportance = "H"
                Else
                    strImportance = "ג"
                End If
            End If
            Select Case strImportance
            Case "H", "ג"
                impH = " checked"
            Case "M", "ב"
                impM = " checked"
            Case "L", "נ"
                impL = " checked"
            Case "NA", "לי"
                impNA = " checked"
            End Select
        End If
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1>" & vbCrLf & "<TR>" & vbCrLf
        If pLanguage = "eng" Then
            tmpStr = tmpStr & "<TD><INPUT type=radio id=importance" & pField_Id & " name=importance" & pField_Id & " value='H' " & impH & "></TD><Td class='form' valign=top>High</Td><td width=5>&nbsp;</td><TD><INPUT type=radio id=importance" & pField_Id & " name=importance" & pField_Id & " value='M' " & impM & "></TD><Td class='form' valign=top>Medium</Td><td width=5>&nbsp;</td><TD><INPUT type=radio id=importance" & pField_Id & " name=importance" & pField_Id & " value='L' " & impL & "></TD><Td class='form' valign=top>Low</Td><td width=5>&nbsp;</td><TD><INPUT type=radio id=importance" & pField_Id & " name=importance" & pField_Id & " value='NA' " & impNA & "></TD><Td class='form' valign=top>Not Applicable</Td><td width=10>&nbsp;</td><Td class='form' valign=top>" & pField_Title & "&nbsp;</Td>"
        End If
        If pLanguage = "heb" Then
            fori = 1
            fotto = CInt(pField_Scale)
            forstep = 1
        Else
            fori = CInt(pField_Scale)
            fotto = 1
            forstep = -1
        End If
         For i = fori To fotto Step forstep
            If CStr(i) = strValue Then
                isChecked = " checked"
            Else
                isChecked = ""
            End If
            If pLanguage = "eng" Then
                tmpStr = tmpStr & "<TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value=" & i & " " & isChecked & "></TD><Td class='form' valign=top>" & i & "</Td><td width=5>&nbsp;</td>" & vbCrLf
            Else
                tmpStr = tmpStr & "<td width=5>&nbsp;</td><Td class='form' valign=top>" & i & "</Td><TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value=" & i & " " & isChecked & "></TD>" & vbCrLf
            End If
        Next
        If pLanguage = "heb" Then
            tmpStr = tmpStr & "<Td class='form' valign=top>&nbsp;" & pField_Title & "</Td><td width=10>&nbsp;</td><Td class='form' valign=top>יטנוולר</Td><TD><INPUT type=radio id=importance" & pField_Id & " name=importance" & pField_Id & " value='לי' " & impNA & "></TD><td width=5>&nbsp;</td><Td class='form' valign=top>ךומנ</Td><TD><INPUT type=radio id=importance" & pField_Id & " name=importance" & pField_Id & " value='נ' " & impL & "></TD><td width=5>&nbsp;</td><Td class='form' valign=top>ינוניב</Td><TD><INPUT type=radio id=importance" & pField_Id & " name=importance" & pField_Id & " value='ב' " & impM & "></TD><td width=5>&nbsp;</td><Td class='form' valign=top>הובג</Td><TD><INPUT type=radio id=importance" & pField_Id & " name=importance" & pField_Id & " value='ג' " & impH & "></TD>"
        End If
        tmpStr = tmpStr & "</TR></TABLE>"
        GetFormField = tmpStr
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
            isChecked = Day(Date) & "/" & Month(Date) & "/" & Year(Date)
        End If
        If pLanguage = "eng" Then
            GetFormField = "<INPUT type=text dir=ltr class=passw maxlength=10 id=field" & pField_Id & " name=field" & pField_Id & " size=10 value='" & isChecked & "' onclick='return popupcal(this);' readonly>&nbsp;<INPUT type=image src='" & pPathCalImage & "images/calend.gif' id=imgcal name=imgcal border=0 onclick='return popupcal(this.form.elements(""field" & pField_Id & """));'>"
        Else
            GetFormField = "<INPUT type=image src='" & pPathCalImage & "images/calend.gif' id=imgcal name=imgcal border=0 onclick='return popupcal(this.form.elements(""field" & pField_Id & """));'>&nbsp;<INPUT type=text dir=ltr class=passw maxlength=10 id=field" & pField_Id & " name=field" & pField_Id & " size=10 value='" & isChecked & "' onclick='return popupcal(this);' readonly>"
        End If
    Case "7" 'number
        GetFormField = "<INPUT type=text dir=ltr class='texts' maxlength=" & pField_Size & " id=field" & pField_Id & " name=field" & pField_Id & " value='" & strValue & "'>"
    Case "8" ' radio - select
         tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1>" & vbCrLf
        Set rsSelect = conn.Execute("Select * from Form_select where Field_id=" & pField_Id & " Order by Id")
            Do While Not rsSelect.EOF
                selValue = rsSelect("Field_Value")
                If selValue = strValue Then
                    isChecked = " checked"
                Else
                    isChecked = ""
                End If
                tmpStr = tmpStr & "<TR>" & vbCrLf
                If pLanguage = "eng" Then
                    tmpStr = tmpStr & "<TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value='" & selValue & "' " & isChecked & "></TD><Td class='form' valign=top>" & selValue & "</Td><td width=5>&nbsp;</td>" & vbCrLf
                Else
                    tmpStr = tmpStr & "<td width=5>&nbsp;</td><Td class='form' dir=rtl valign=top>" & selValue & "</Td><TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value='" & selValue & "' " & isChecked & "></TD>" & vbCrLf
                End If
                tmpStr = tmpStr & "</TR>" & vbCrLf
            rsSelect.MoveNext
            Loop
        Set rsSelect = Nothing
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
                tmpStr = tmpStr & "<TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value=" & i & " " & isChecked & "></TD><Td class='form' valign=top>" & i & "</Td><td width=5>&nbsp;</td>" & vbCrLf
            Else
                tmpStr = tmpStr & "<td width=5>&nbsp;</td><Td class='form' valign=top>" & i & "</Td><TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value=" & i & " " & isChecked & "></TD>" & vbCrLf
            End If
            tmpStr = tmpStr & "</TR>" & vbCrLf
        Next
        tmpStr = tmpStr & "</TABLE>"
        GetFormField = tmpStr
    'Case "10" - subject - can't use here
    Case "11" ' radio - select - vertical
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=rtl>" & vbCrLf
        tmpStr = tmpStr & "<TR>" & vbCrLf
        Set rsSelect = conn.Execute("Select * from Form_select where Field_id=" & pField_Id & " Order by Id")
            Do While Not rsSelect.EOF
                selValue = rsSelect("Field_Value")
                If selValue = strValue Then
                    isChecked = " checked"
                Else
                    isChecked = ""
                End If
                If pLanguage = "eng" Then
                    tmpStr = tmpStr & "<TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value='" & selValue & "' " & isChecked & "></TD><Td class='form' valign=top>" & selValue & "</Td><td width=5>&nbsp;</td>" & vbCrLf
                Else
                    tmpStr = tmpStr & "<td width=5>&nbsp;</td><TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value='" & selValue & "' " & isChecked & "></TD><Td class='form' dir=rtl valign=top>" & selValue & "</Td>" & vbCrLf
                End If
            rsSelect.MoveNext
            Loop
        Set rsSelect = Nothing
        tmpStr = tmpStr & "</TR>" & vbCrLf
        tmpStr = tmpStr & "</TABLE>"
        GetFormField = tmpStr
      Case "12" 'scale - vertical
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=rtl>" & vbCrLf
        tmpStr = tmpStr & "<TR>" & vbCrLf
        For i = 1 To CInt(pField_Scale)
            If CStr(i) = strValue Then
                isChecked = " checked"
            Else
                isChecked = ""
            End If
            If pLanguage = "eng" Then
                tmpStr = tmpStr & "<TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value=" & i & " " & isChecked & "></TD><Td class='form' valign=top>" & i & "</Td><td width=5>&nbsp;</td>" & vbCrLf
            Else
                tmpStr = tmpStr & "<td width=5>&nbsp;</td><TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value=" & i & " " & isChecked & "></TD><Td class='form' valign=top>" & i & "</Td>" & vbCrLf
            End If
        Next
        tmpStr = tmpStr & "</TR>" & vbCrLf
        tmpStr = tmpStr & "</TABLE>"
        GetFormField = tmpStr
     End Select
End Function

Public Function GetFieldGraph(pProduct_Id, pField_Id, pField_Title, pField_Type, pField_Scale, pLanguage) As String
Dim tmpStr As String, selValue As String, colind As Integer, selAll As Long, wd As Integer, i As Integer
Dim td_align As String
Dim ScQuant As Integer
Dim rsSelect As Recordset
Dim col(9) As String
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
        Set rsSelect = GetRecordSet("Select * from FIELD_STATISTICS where Product_Id = " & pProduct_Id & " and  Field_id=" & pField_Id & " Order by Field_Order " & string_sort)
        
        selAll = 0
        tmpStr = "<tr><td bgcolor=#D8E0E6 class='form' align=" & td_align & " nowrap valign=top colspan=" & rsSelect.RecordCount & " dir=rtl>&nbsp;" & pField_Title & "&nbsp;</td></tr>" & vbCrLf
        If Not rsSelect.EOF Then
            tmpStr = tmpStr & "<tr>" & vbCrLf
            Do While Not rsSelect.EOF
                selAll = selAll + CInt(rsSelect("FIELD_QUANTITY"))
            rsSelect.MoveNext
            Loop
            rsSelect.MoveFirst
            Do While Not rsSelect.EOF
                If selAll <> 0 Then
                    wd = CInt((CInt(rsSelect("FIELD_QUANTITY")) / selAll) * 100)
                    If rsSelect("FIELD_QUANTITY") <> 0 Then
                        str_FIELD_QUANTITY = "<br><span style=color:blue;font-size:10pt>" & rsSelect("FIELD_QUANTITY") & "</span>"
                    Else
                        str_FIELD_QUANTITY = ""
                    End If
                Else
                    wd = 0
                End If
                tmpStr = tmpStr & "<td align=center valign=bottom height=120 width=60>" & vbCrLf
                tmpStr = tmpStr & "<TABLE align=center BORDER=0 CELLSPACING=1 CELLPADDING=1 width=20>" & vbCrLf
                tmpStr = tmpStr & "<tr><td align=center><font size=2><font color=red>" & wd & "</font>%</font>" & str_FIELD_QUANTITY & "</td></tr>" & vbCrLf
                tmpStr = tmpStr & "<tr><td align=center valign=bottom><table border=0 cellspacing=0 cellpadding=0>" & vbCrLf
                tmpStr = tmpStr & "<tr><td width=20 bgcolor=" & col(colind Mod 9) & " height=" & wd & "></td></tr>" & vbCrLf
                tmpStr = tmpStr & "</table></td></tr></table></td>" & vbCrLf
            rsSelect.MoveNext
            colind = colind + 1
            Loop
            tmpStr = tmpStr & "</tr><tr>" & vbCrLf
            rsSelect.MoveFirst
            Do While Not rsSelect.EOF
                tmpStr = tmpStr & "<td align=center bgcolor=#D8E0E6>&nbsp;<font size=2 dir=rtl>" & rsSelect("FIELD_SELECT") & "</font>&nbsp;</td>"
            rsSelect.MoveNext
            Loop
            tmpStr = tmpStr & "</tr>"
        End If  'not rsSelect.EOF
                            
        Set rsSelect = Nothing
       
        GetFieldGraph = tmpStr
    Case "9", "12" 'radio
        Set rsSelect = GetRecordSet("SELECT * FROM QUESTION_VALUES_VIEW WHERE Product_Id = " & pProduct_Id & " and Field_Id=" & pField_Id & " and FIELD_VALUE is not null")
        selAll = rsSelect.RecordCount
            
        tmpStr = "<tr><td bgcolor=#D8E0E6 class='form'  align=" & td_align & " nowrap valign=top colspan=" & pField_Scale & " dir=rtl>&nbsp;" & pField_Title & "&nbsp;</td></tr>" & vbCrLf
        tmpStr = tmpStr & "<tr>" & vbCrLf
        
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
                str_FIELD_QUANTITY = "<br><span style=color:blue;font-size:10pt>" & ScQuant & "</span>"
            Else
                wd = 0
                str_FIELD_QUANTITY = ""
            End If
            tmpStr = tmpStr & "<td align=center valign=bottom height=120 width=60>" & vbCrLf
            tmpStr = tmpStr & "<TABLE align=center BORDER=0 CELLSPACING=1 CELLPADDING=1 width=20>" & vbCrLf
            tmpStr = tmpStr & "<tr><td align=center><font size=2><font color=red>" & wd & "</font>%</font>" & str_FIELD_QUANTITY & "</td></tr>" & vbCrLf
            tmpStr = tmpStr & "<tr><td align=center valign=bottom><table border=0 cellspacing=0 cellpadding=0>" & vbCrLf
            tmpStr = tmpStr & "<tr><td width=20 bgcolor=" & col(colind Mod 9) & " height=" & wd & "></td></tr>" & vbCrLf
            tmpStr = tmpStr & "</table></td></tr></table></td>" & vbCrLf
        colind = colind + 1
        Next
        Set rsSelect = Nothing
        tmpStr = tmpStr & "</tr><tr>" & vbCrLf
        For i = fori To fotto Step forstep
            tmpStr = tmpStr & "<td align=center bgcolor=#D8E0E6>&nbsp;<font size=2>" & i & "</font>&nbsp;</td>"
        Next
        tmpStr = tmpStr & "</tr>"
        
        GetFieldGraph = tmpStr
    Case "5" 'check
        Set rsSelect = GetRecordSet("SELECT * FROM QUESTION_VALUES_VIEW WHERE Product_Id = " & pProduct_Id & " and Field_Id=" & pField_Id)
        selAll = rsSelect.RecordCount
        tmpStr = "<tr><td bgcolor=#D8E0E6 class='form'  align=" & td_align & " nowrap valign=top colspan=2 dir=rtl>&nbsp;" & pField_Title & "&nbsp;</td></tr>" & vbCrLf
        tmpStr = tmpStr & "<tr>" & vbCrLf
        rsSelect.Filter = "FIELD_VALUE='on'"
        ScQuant = rsSelect.RecordCount
        Set rsSelect = Nothing
        If ScQuant <> 0 Then
            wd = CInt((ScQuant / selAll) * 100)
            str_FIELD_QUANTITY = "<br><span style=color:blue;font-size:10pt>" & ScQuant & "</span>"
        Else
            wd = 0
            str_FIELD_QUANTITY = ""
        End If
        tmpStr = tmpStr & "<td align=center valign=bottom height=120 width=60>" & vbCrLf
        tmpStr = tmpStr & "<TABLE align=center BORDER=0 CELLSPACING=1 CELLPADDING=1 width=20>" & vbCrLf
        tmpStr = tmpStr & "<tr><td align=center><font size=2><font color=red>" & wd & "</font>%</font>" & str_FIELD_QUANTITY & "</td></tr>" & vbCrLf
        tmpStr = tmpStr & "<tr><td align=center valign=bottom><table border=0 cellspacing=0 cellpadding=0>" & vbCrLf
        tmpStr = tmpStr & "<tr><td width=20 bgcolor=" & col(colind Mod 9) & " height=" & wd & "></td></tr>" & vbCrLf
        tmpStr = tmpStr & "</table></td></tr></table></td>" & vbCrLf
        
        ScQuant = selAll - ScQuant
        If ScQuant <> 0 Then
            wd = CInt((ScQuant / selAll) * 100)
            str_FIELD_QUANTITY = "<br><span style=color:blue;font-size:10pt>" & ScQuant & "</span>"
        Else
            wd = 0
            str_FIELD_QUANTITY = ""
        End If
        colind = colind + 1
            tmpStr = tmpStr & "<td align=center valign=bottom height=120 width=60>" & vbCrLf
        tmpStr = tmpStr & "<TABLE align=center BORDER=0 CELLSPACING=1 CELLPADDING=1 width=20>" & vbCrLf
        tmpStr = tmpStr & "<tr><td align=center><font size=2><font color=red>" & wd & "</font>%</font>" & str_FIELD_QUANTITY & "</td></tr>" & vbCrLf
        tmpStr = tmpStr & "<tr><td align=center valign=bottom><table border=0 cellspacing=0 cellpadding=0>" & vbCrLf
        tmpStr = tmpStr & "<tr><td width=20 bgcolor=" & col(colind Mod 9) & " height=" & wd & "></td></tr>" & vbCrLf
        tmpStr = tmpStr & "</table></td></tr></table></td>" & vbCrLf
        tmpStr = tmpStr & "</tr><tr><td align=center nowrap bgcolor=#D8E0E6>&nbsp;<font size=2>כן</font>&nbsp;</td><td align=center nowrap bgcolor=#D8E0E6>&nbsp;<font size=2>לא</font>&nbsp;</td></tr>"
         
        GetFieldGraph = tmpStr
    Case "6" 'date
        GetFieldGraph = ""
    Case "7" 'number
        GetFieldGraph = ""
        
        
    End Select
End Function
Public Function reverse(str As Variant) As Variant
   ' A function which reverses one line of hebrew text
   ' RETURNS: The reversed hebrew string (Null if string is Null).
   
   On Error Resume Next 'GoTo ErrElse
   
   Dim revword As Variant
   Dim engword As Variant
   Dim i As Integer
   Dim letter As Variant
   Dim num As Integer
   Dim eng As Variant
   Dim heb As Variant
   Dim pos As Integer
   Dim newword As Variant
   Dim postcheck As Variant
   
   If str = "" Or IsNull(str) Then
      reverse = ""
      Exit Function
   End If
   
   revword = ""
   engword = ""
   heb = " ,./':-()%" + Chr(34)
   eng = ".,?!'()" + Chr(34)
   postcheck = "."
   pos = 0
   i = 1
   
   Do While i <= Len(str)
      letter = Mid(str, i, 1)
      num = Asc(letter)
      If (num >= 224 And num <= 250) Then
         Exit Do
      End If
      i = i + 1
   Loop
   
   If i <= Len(str) Then
          
   For i = 1 To Len(str)
      letter = Mid(str, i, 1)
      num = Asc(letter)
      If (num >= 224 And num <= 250) Or (InStr(1, heb, letter) <> 0 And pos = 0) Then
'hebrew
                  If letter = "(" Then
                    letter = ")"
                  ElseIf letter = ")" Then
                    letter = "("
                  End If
                  If pos > 0 Then
                    revword = " " & revword
                  End If
                  
                  revword = letter & revword
                  pos = 0
      Else
'english
                  If InStr(1, eng, letter) = 0 Or (InStr(1, postcheck, letter) <> 0 And Mid(str, i + 1, 1) <> " " And Mid(str, i + 1, 1) <> "") Then
                    revword = Left(revword, pos) & letter & Mid(revword, pos + 1)
                    pos = pos + 1
                  Else
'hebrew
                    If letter = "(" Then
                      letter = ")"
                    ElseIf letter = ")" Then
                      letter = "("
                    End If
                    revword = letter & revword
                    pos = 0
                  End If
      End If
   Next i
   reverse = revword
   Else
      reverse = str
   End If

Exit_Label:
   Exit Function
ErrElse:
   Err.Raise Err.Number, Err.Source, "Reverse(str) Function COM Error:" & vbCrLf & Err.Description & vbCrLf & "Contact Michael Kalinka for details."
   Resume Exit_Label
End Function


Public Function hebrew(word, width) As Variant
    Dim place, newstring, newword, pos, pos1, insstr, spflag, entpos
    If word = "" Or IsNull(word) Then
      hebrew = ""
      Exit Function
    End If
         width = width + 1
        word = Trim(word)
        spflag = False
        For pos = 1 To Len(word)
          If Mid(word, pos, 1) <> " " Then
             newword = newword & Mid(word, pos, 1)
             spflag = False
          Else
            If spflag = False Then
                newword = newword & " "
                spflag = True
            End If
          End If
        Next
        word = newword

    newstring = ""
    Do While Len(word) > 0
                If width < Len(word) Then
              place = width
          Do While Mid(word, place, 1) <> " " And place > 1
            place = place - 1
          Loop
          If place = 1 Then
                     place = width
                     Do While place <= Len(word)
                       If Mid(word, place, 1) = " " Then Exit Do
                       place = place + 1
                     Loop
                  End If

                Else
                  place = Len(word) + 1
                End If

                entpos = InStr(word, Chr(13) & Chr(10))
                If entpos <> 0 And entpos <= place Then
                  word = Left(word, entpos - 1) & " " & Mid(word, entpos + 1)
                  place = entpos
                End If

                newword = reverse(Left(word, place - 1))
                pos = 1
                Do While pos <> 0
                  pos = InStr(pos, newword, "http://", 1)
                  If pos <> 0 Then
                    pos1 = InStr(pos, newword, " ", 1) - 1
                    If pos1 < 0 Then pos1 = Len(newword)

                    v_str = Mid(newword, pos, pos1 - pos + 1)
                    v_pos = InStr(8, v_str, "/", 1)
                    'if v_pos<>0 then
                    '  v_str0=left(v_str,v_pos-1)+"..."
                    'else
                      v_str0 = v_str
                    'end if

                    insstr = "<a href='" + v_str + "' target=_blank>" + v_str0 + "</a>"
                    newword = Left(newword, pos - 1) + insstr + Mid(newword, pos1 + 1)
                    pos = pos + Len(insstr)
                  End If
                Loop

                pos = 1
                Do While pos <> 0
                  pos = InStr(pos, newword, "www.", 1)
                  If pos <> 0 Then
                    spflag = True
                    If pos >= 7 Then
                      If Mid(newword, pos - 7, 7) = "http://" Then
                        pos = pos + 1
                        spflag = False
                      End If
                    End If
                    If spflag = True Then
                      pos1 = InStr(pos, newword, " ", 1) - 1
                      If pos1 < 0 Then pos1 = Len(newword)

                      v_str = Mid(newword, pos, pos1 - pos + 1)
                      v_pos = InStr(1, v_str, "/", 1)
                      'if v_pos<>0 then
                      '  v_str0=left(v_str,v_pos-1)+"..."
                      'else
                       v_str0 = v_str
                      'end if
                
                      insstr = "<a href='http://" + v_str + "' target=_blank>" + v_str0 + "</a>"
                      newword = Left(newword, pos - 1) + insstr + Mid(newword, pos1 + 1)
                      pos = pos + Len(insstr)
                    End If
                  End If
                Loop

                pos = 1
                Do While pos <> 0
                  pos = InStr(pos, newword, "@", 1)
                  If pos <> 0 Then
                    Do While pos > 0
                      If Mid(newword, pos, 1) = " " Then
                        Exit Do
                      End If
                      pos = pos - 1
                    Loop
                    pos = pos + 1

                    pos1 = InStr(pos, newword, " ", 1) - 1
                    If pos1 <= 0 Then pos1 = Len(newword)

                    insstr = "<a href='mailto:" + Mid(newword, pos, pos1 - pos + 1) + "'>" + Mid(newword, pos, pos1 - pos + 1) + "</a>"
                    newword = Left(newword, pos - 1) + insstr + Mid(newword, pos1 + 1)
                    pos = pos + Len(insstr)
                  End If
                Loop

        newstring = newstring & newword
        word = Right(word, Len(word) - place + 1)
                word = Trim(word)
                If Len(word) > 0 Then newstring = newstring & "<br>"
              Loop
              newstring = newstring & reverse(word)
    hebrew = newstring
    
End Function


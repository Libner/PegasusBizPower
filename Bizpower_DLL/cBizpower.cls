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
    conn.Close
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
        GetFormField = "<INPUT type='text' dir=" & pField_Align & " maxlength=" & pField_Size & " size=" & Field_Size & " id=field" & pField_Id & " name=field" & pField_Id & " value=""" & Replace(strValue, Chr(34), "''") & """>"
    Case "2" 'textarea
        GetFormField = "<textarea dir=" & pField_Align & " class='txt' rows=5 cols=60 id=field" & pField_Id & " name=field" & pField_Id & ">" & strValue & "</textarea>"
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
                selValue = Replace(selValue, Chr(34), "''")
                If selValue = strValue Then
                    isSelected = " selected"
                Else
                    isSelected = ""
                End If
                tmpStr = tmpStr & "<option value=""" & selValue & """ " & isSelected & ">" & selValue & "</option>" & vbCrLf
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
        If pLanguage = "eng" Then
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=ltr>" & vbCrLf
        Else
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=rtl>" & vbCrLf
        End If
        tmpStr = tmpStr & "<TR>" & vbCrLf
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
            isChecked = vbNullString
        End If
        If pLanguage = "eng" Then
            GetFormField = "<INPUT type=text dir=ltr class=passw maxlength=10 id=field" & pField_Id & " name=field" & pField_Id & " size=10 value='" & isChecked & "' onclick='return popupcal(this);return false;' readonly>&nbsp;<Img class='hand' src='" & pPathCalImage & "images/calend.gif' id=imgcal name=imgcal border=0 onclick='return popupcal(window.document.all(""field" & pField_Id & """));return false;'>"
        Else
            GetFormField = "<Img class='hand' src='" & pPathCalImage & "images/calend.gif' id=imgcal name=imgcal border=0 onclick='return popupcal(window.document.all(""field" & pField_Id & """));return false;'>&nbsp;<INPUT type=text dir=ltr class=passw maxlength=10 id=field" & pField_Id & " name=field" & pField_Id & " size=10 value='" & isChecked & "' onclick='return popupcal(this);return false;' readonly>"
        End If
    Case "7" 'number
        GetFormField = "<INPUT type=text dir=ltr class='texts' maxlength=" & pField_Size & " id=field" & pField_Id & " name=field" & pField_Id & " value=""" & Replace(strValue, Chr(34), "''") & """>"
    Case "8" ' radio - select
      If pLanguage = "eng" Then
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=ltr>" & vbCrLf
        Else
        tmpStr = "<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=1 dir=rtl>" & vbCrLf
        End If
        
        Set rsSelect = conn.Execute("Select * from Form_select where Field_id=" & pField_Id & " Order by Id")
            Do While Not rsSelect.EOF
                selValue = rsSelect("Field_Value")
                selValue = Replace(selValue, Chr(34), "''")
                If selValue = strValue Then
                    isChecked = " checked"
                Else
                    isChecked = ""
                End If
                tmpStr = tmpStr & "<TR>" & vbCrLf
                If pLanguage = "eng" Then
                    tmpStr = tmpStr & "<TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value=""" & selValue & """ " & isChecked & "></TD><Td class='form' valign=top>" & selValue & "</Td><td width=5>&nbsp;</td>" & vbCrLf
                Else
                    tmpStr = tmpStr & "<td width=5>&nbsp;</td><TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value=""" & selValue & """ " & isChecked & "></TD><Td class='form' dir=rtl valign=top>" & selValue & "</td>" & vbCrLf
                End If
                tmpStr = tmpStr & "</TR>" & vbCrLf
            rsSelect.MoveNext
            Loop
        Set rsSelect = Nothing
        
        'If strValue = "" Then
        '    isChecked = " checked"
        'Else
        '    isChecked = ""
        'End If
        'tmpStr = tmpStr & "<TR>" & vbCrLf
        'If pLanguage = "eng" Then
        '    tmpStr = tmpStr & "<TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value='' " & isChecked & "></TD><Td class='form' valign=top>Not relevant</Td><td width=5>&nbsp;</td>" & vbCrLf
        'Else
        '    tmpStr = tmpStr & "<td width=5>&nbsp;</td><TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value='' " & isChecked & "></TD><Td class='form' dir=rtl valign=top>לא רלונטי</td>" & vbCrLf
        'End If
        'tmpStr = tmpStr & "</TR>" & vbCrLf
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
        
        Set rsSelect = conn.Execute("Select * from Form_select where Field_id=" & pField_Id & " Order by Id")
            Do While Not rsSelect.EOF
                selValue = rsSelect("Field_Value")
                selValue = Replace(selValue, Chr(34), "''")
                If selValue = strValue Then
                    isChecked = " checked"
                Else
                    isChecked = ""
                End If
                If pLanguage = "eng" Then
                    tmpStr = tmpStr & "<TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value=""" & selValue & """ " & isChecked & "></TD><Td class='form' valign=top>" & selValue & "</Td><td width=5>&nbsp;</td>" & vbCrLf
                Else
                    tmpStr = tmpStr & "<td width=5>&nbsp;</td><TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value=""" & selValue & """ " & isChecked & "></TD><Td class='form' dir=rtl valign=top>" & selValue & "</Td>" & vbCrLf
                End If
            rsSelect.MoveNext
            Loop
        Set rsSelect = Nothing
        'If strValue = "" Then
        '    isChecked = " checked"
        'Else
        '    isChecked = ""
        'End If
        'If pLanguage = "eng" Then
        '    tmpStr = tmpStr & "<TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value='' " & isChecked & "></TD><Td class='form' valign=top>Not relevant</Td><td width=10>&nbsp;</td>" & vbCrLf
        'Else
        '    tmpStr = tmpStr & "<td width=10>&nbsp;</td><TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value='' " & isChecked & "></TD><Td class='form' dir=rtl valign=top>לא רלונטי</Td>" & vbCrLf
        'End If
        
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
                tmpStr = tmpStr & "<TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value=" & i & " " & isChecked & "></TD><Td class='form' valign=top>" & i & "</Td><td width=5>&nbsp;</td>" & vbCrLf
            Else
                tmpStr = tmpStr & "<td width=5>&nbsp;</td><TD><INPUT type=radio id=field" & pField_Id & " name=field" & pField_Id & " value=" & i & " " & isChecked & "></TD><Td class='form' valign=top>" & i & "</Td>" & vbCrLf
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

Public Function GetFieldGraph(pProduct_Id, pField_Id, pField_Title, pField_Type, pField_Scale, pLanguage, pType_ID) As String
Dim tmpStr As String, selValue As String, colind As Integer, selAll As Long, wd As Integer, i As Integer
Dim td_align As String
Dim td_width As Integer
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
        Set rsSelect = GetRecordSet("Select * from FIELD_STATISTICS where Product_Id = " & pProduct_Id & " and  Field_id=" & pField_Id & " and Type_ID = " & pType_ID & " Order by Field_Order " & string_sort)
        
        selAll = 0
        tmpStr = "<tr><td bgcolor=#DADADA class='form' align=" & td_align & " nowrap valign=top colspan=" & rsSelect.RecordCount & "><b>&nbsp;" & pField_Title & "&nbsp;</td></tr>" & vbCrLf
        If rsSelect.RecordCount <> 0 Then
            td_width = 100 / rsSelect.RecordCount
        End If
        
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
                        str_FIELD_QUANTITY = " &nbsp;<span style=color:blue;font-size:10pt>(" & rsSelect("FIELD_QUANTITY") & ")</span>"
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
            rsSelect.MoveNext
            colind = colind + 1
            Loop
            tmpStr = tmpStr & "</tr><tr>" & vbCrLf
            rsSelect.MoveFirst
            Do While Not rsSelect.EOF
                tmpStr = tmpStr & "<td align=center bgcolor=#DADADA>&nbsp;<font size=2><b>" & rsSelect("FIELD_SELECT") & "</font>&nbsp;</td>"
            rsSelect.MoveNext
            Loop
            tmpStr = tmpStr & "</tr>"
        End If  'not rsSelect.EOF
                            
        Set rsSelect = Nothing
       
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

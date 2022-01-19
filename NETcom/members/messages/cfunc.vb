Namespace bizpower
    Public Class cfunc
        Protected WithEvents myConnection As System.Data.SqlClient.SqlConnection
        Dim myReader As System.Data.SqlClient.SqlDataReader
        Protected WithEvents myCommand As System.Data.SqlClient.SqlCommand

        Function sFix(ByVal myString As Object) As String
            Dim quote, newQuote, temp As String
            If Not IsDBNull(myString) Then
                temp = CStr(myString)
            Else
                Return ""
                Exit Function
            End If

            If temp <> "" Then
                quote = "'"
                newQuote = "''"
                If InStr(myString, "'") Then
                    temp = Replace(temp, quote, newQuote)
                End If
                temp = Replace(temp, "'", "'+char(39)+'")
                temp = Replace(temp, Chr(150), Chr(45)) ' "-"
            End If
            Return temp

        End Function

        Function GetLanguageScript(ByVal lang_id As String, ByVal page_id As String) As String
            Dim strHtml, sqlstr, word_id, word_str As String
            strHtml = "<SCRIPT>" & vbCrLf
            strHtml += "var arrElem;" & vbCrLf
            If lang_id <> "" And page_id <> "" Then
                sqlstr = "Select word_id, word From translate Where lang_id=" & lang_id & " And page_id=" & page_id & "  Order By word_id"
                myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
                myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
                myConnection.Open()
                myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)

                While myReader.Read()
                    word_id = Trim(myReader("word_id"))
                    word_str = Trim(myReader("word"))
                    strHtml += "arrElem = window.document.getElementsByName('word" & word_id & "');" & vbCrLf
                    strHtml += "if(arrElem != null && arrElem.item(0) != null)" & vbCrLf
                    strHtml += "{		if(arrElem.item(0).title)" & vbCrLf
                    strHtml += "{" & vbCrLf
                    strHtml += "if(arrElem.length != null)" & vbCrLf
                    strHtml += "{  for(i=0;i<arrElem.length;i++)" & vbCrLf
                    strHtml += "		arrElem.item(i).title = """ & word_str & """ ;" & vbCrLf
                    strHtml += "}" & vbCrLf
                    strHtml += "else  arrElem.item(0).title = """ & word_str & """;" & vbCrLf
                    strHtml += "}" & vbCrLf
                    strHtml += "else if(arrElem.item(0).innerHTML)" & vbCrLf
                    strHtml += "arrElem.item(0).innerHTML = """ & word_str & """;" & vbCrLf
                    strHtml += "}" & vbCrLf
                End While
                myReader.Close()
            Else
                strHtml += "alert('Check language id=" & lang_id & " or page id=" & page_id & "');" & vbCrLf
            End If
            strHtml += "</SCRIPT>" & vbCrLf

            Return strHtml

        End Function

        Function vFix(ByVal word As Object) As String
            Dim temp As String
            If Not IsDBNull(word) Then
                temp = CStr(word)
            Else
                Return ""
                Exit Function
            End If
            temp = Replace(temp, "<", "&lt;")
            temp = Replace(temp, ">", "&gt;")
            temp = Replace(temp, Chr(34), "&quot;")
            temp = Replace(temp, Chr(39), "&#39;")

            Return temp

        End Function

        Function Encode(ByVal sIn)
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

        Function Decode(ByVal sIn)
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

        Function dbNullFix(ByVal objVal As Object) As String
            If IsDBNull(objVal) Then
                dbNullFix = ""
            Else
                dbNullFix = Trim(objVal)
            End If
        End Function

        Public Function fixNumeric(ByVal objVal As Object) As Double
            If IsNumeric(objVal) Then
                Return CDbl(objVal)
            Else
                Return 0
            End If
        End Function

    End Class
End Namespace


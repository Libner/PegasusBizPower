Namespace bizpower
    Public Class cfunc
        Protected WithEvents myConnection As System.Data.SqlClient.SqlConnection
        Dim myReader As System.Data.SqlClient.SqlDataReader
        Protected WithEvents myCommand As System.Data.SqlClient.SqlCommand
        Public PegasusSiteDBName As String = ConfigurationSettings.AppSettings("PegasusSiteDBName")
        Function dbNullDate(ByVal objVal As Object, ByVal formatString As String) As String
            Dim strDate As String = ""
            If IsDBNull(objVal) Then
                strDate = ""
            Else
                If IsDate(objVal) Then
                    If formatString <> "" Then
                        strDate = CDate(objVal).ToString(formatString)
                    Else
                        strDate = CDate(objVal)
                    End If
                End If
            End If
            dbNullDate = strDate
        End Function
        Function dbNullBool(ByVal objVal As Object) As Boolean
            Try

                If IsNothing(objVal) Then
                    dbNullBool = False
                Else
                    If IsDBNull(objVal) Then
                        dbNullBool = False
                    Else
                        dbNullBool = CBool(objVal)
                    End If
                End If
            Catch ex As Exception
                dbNullBool = False
            End Try
        End Function

        Function RandomPassword(ByVal PasswordLength As Integer, ByVal NumberOfLowerCases As Integer, ByVal NumberOfUpperCases As Integer, ByVal NumberOfNumbers As Integer) As String

            Dim lNumberOfLowerCases = NumberOfLowerCases
            Dim lNumberOfUpperCases = NumberOfUpperCases
            Dim lNumberOfNumbers = NumberOfNumbers
            Dim l, j

            Dim ReturnedPassword As String = ""
            If PasswordLength < 3 Then
                Exit Function
            End If

            'Get the number of digits for each type of characters
            Randomize()
            'lNumberOfLowerCases = CInt((PasswordLength - 3) * Rnd()) + 1
            'lNumberOfUpperCases = CInt((PasswordLength - lNumberOfLowerCases - 2) * Rnd()) + 1
            '' lNumberOfUpperCases = 0
            'lNumberOfNumbers = PasswordLength - lNumberOfLowerCases - lNumberOfUpperCases


            ReturnedPassword = ""
            For l = 1 To PasswordLength
                Randomize()
                j = CInt(2 * Rnd() + 1)
                Select Case j
                    Case 1 'Lower Case
                        If lNumberOfLowerCases > 0 Then
                            ReturnedPassword = ReturnedPassword & Chr(CInt(25 * Rnd()) + 97)
                            lNumberOfLowerCases = lNumberOfLowerCases - 1
                        Else
                            l = l - 1 'Re-do the loop
                        End If
                    Case 2 'Upper Case
                        If lNumberOfUpperCases > 0 Then
                            ReturnedPassword = ReturnedPassword & Chr(CInt(25 * Rnd()) + 65)
                            lNumberOfUpperCases = lNumberOfUpperCases - 1
                        Else
                            l = l - 1 'Re-do the loop
                        End If
                    Case 3 'Number
                        If lNumberOfNumbers > 0 Then
                            ReturnedPassword = ReturnedPassword & CInt(9 * Rnd())
                            lNumberOfNumbers = lNumberOfNumbers - 1
                        Else
                            l = l - 1 'Re-do the loop
                        End If
                End Select
                For j = 1 To 100
                    'Give the seed some time
                Next
            Next

            RandomPassword = ReturnedPassword
        End Function
        Public Function ShortWeekDay(ByVal WeekDay As String) As String

            'Current.Response.Write(currDate)
            Return ShortWeekDay
        End Function
        Public Function fncCurrentDate() As Date
            Dim currDate As Date = Now
            If (currDate.Hour >= 0) And (currDate.Hour < 2) Then
                currDate = currDate.AddDays(-1)
            End If
            'Current.Response.Write(currDate)
            Return currDate
        End Function
        Function NumFeedBackPeople(ByVal DepartureId As String) As String
            If DepartureId <> "" Then
                Dim myReader As System.Data.SqlClient.SqlDataReader
                Dim myConnection As New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))
                myConnection.Open()
                Dim sqlstr As String
                Dim result As Integer = 0
                sqlstr = " SELECT sum(cast(FORM_VALUE.FIELD_VALUE as integer)) as NumFeedBackPeople" & _
                " FROM   CONTACTS INNER JOIN  APPEALS ON CONTACTS.CONTACT_ID = APPEALS.CONTACT_ID INNER JOIN" & _
                " FORM_VALUE ON APPEALS.APPEAL_ID = FORM_VALUE.APPEAL_ID inner join " & PegasusSiteDBName & ".dbo.FeedBack_Form FF on FF.Contact_Id=APPEALS.Contact_Id and FF.Departure_Id=APPEALS.Departure_Id " & _
                " WHERE (APPEALS.Departure_Id = " & DepartureId & ") and field_id=40660  and IsNumeric(FORM_VALUE.FIELD_VALUE)=1 and FeedBack_Status = '2'" & _
                " and exists (select FIELD_ID from FORM_VALUE as confirm where APPEALS.APPEAL_ID = confirm.APPEAL_ID and field_id=40661 and isnull(FIELD_VALUE,'')<>'on')"

                Dim myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
                myReader = myCommand.ExecuteReader()
                While myReader.Read()
                    If Not IsDBNull(myReader("NumFeedBackPeople")) Then
                        result = myReader("NumFeedBackPeople")
                    Else
                        result = 0
                    End If

                End While
                myConnection.Close()

                Return result


            End If

        End Function
        '  Function CheckData(ByVal ContactId As Integer, ByVal CRM_Country As String, ByVal DD As Date) As String
        '   Dim result As String = ""
        '    Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        '    Dim cmdSelect As New SqlClient.SqlCommand("select max(Appeal_Id) as s from Appeals where  QUESTIONS_ID=16470  and APPEAL_STATUS not in(3,5) and Contact_id=" & ContactId & " AND DateDiff(d,A.appeal_date, convert(smalldatetime,''' + " & dateEnd & " + ''',101)) >= 0 ' and  appeal_CountryId in (" & CRM_Country & ")", myConnection)
        '    cmdSelect.CommandType = CommandType.Text
        '    myConnection.Open()
        '
        '           myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
        '          While myReader.Read()
        '             If Not IsDBNull(myReader("s")) Then
        '
        '                   result = 2
        '
        '               Else
        '                  result = "0"
        '             End If
        '        End While
        '       myConnection.Close()
        '
        '
        '           Return result
        '
        '       End Function
        Function Avrg_Call(ByVal pMonth As String, ByVal pYear As String) As String
            Dim result As String = ""
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            Dim cmdSelect As New SqlClient.SqlCommand("SELECT Avrg_Call from DimData_Monthly where Dim_Month =" & pMonth & " and DimYear=" & pYear, myConnection)
            cmdSelect.CommandType = CommandType.Text
            myConnection.Open()

            myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
            While myReader.Read()
                If Not IsDBNull(myReader("Avrg_Call")) Then

                    result = myReader("Avrg_Call")

                Else
                    result = "0"
                End If




            End While
            'If result = "" Then
            '    result = "הכל"
            'End If
            myConnection.Close()


            Return result
        End Function
        Function PermissionsByDepartmentId(ByVal UID As Integer) As String
            Dim result As String = ""
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            Dim cmdSelect As New SqlClient.SqlCommand("SELECT IsNULL(PermissionsByDepartmentId, 0) as PermissionsByDepartmentId  FROM Users where User_id=" & UID, myConnection)
            cmdSelect.CommandType = CommandType.Text
            myConnection.Open()
            result = cmdSelect.ExecuteScalar
            myConnection.close()
            Return result
        End Function
        Function YearStatus(ByVal pYear As String) As String
            Dim result As String = ""
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            Dim cmdSelect As New SqlClient.SqlCommand("SELECT Year_Status from YearsStatus where Year =" & pYear, myConnection)
            cmdSelect.CommandType = CommandType.Text
            myConnection.Open()

            myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
            While myReader.Read()
                If Not IsDBNull(myReader("Year_Status")) Then

                    result = myReader("Year_Status")

                Else
                    result = "0"
                End If




            End While
            If result = "" Or result = 0 Then
                result = "עמוד לא תקין"
            Else
                result = "עמוד תקין"
            End If
            myConnection.Close()


            Return result
        End Function
        Function Avrg_CallProc(ByVal pMonth As String, ByVal pYear As String) As String
            Dim result As String = ""
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            Dim cmdSelect As New SqlClient.SqlCommand("SELECT Avrg_CallProc from DimData_Monthly where Dim_Month =" & pMonth & " and DimYear=" & pYear, myConnection)
            cmdSelect.CommandType = CommandType.Text
            myConnection.Open()

            myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
            While myReader.Read()
                If Not IsDBNull(myReader("Avrg_CallProc")) Then

                    result = myReader("Avrg_CallProc")

                Else
                    result = "0"
                End If




            End While
            'If result = "" Then
            '    result = "הכל"
            'End If
            myConnection.Close()


            Return result
        End Function
        Function GetVouchers_ProviderStatus(ByVal ValueStr As String) As String
            Dim result As String = ""
            Dim ss As String
            ss = ValueStr   'Departure_Id
            '  Response.Write("ss=" & ss)
            '  Response.End()
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            If ss <> "" Then

                Dim cmdSelect As New SqlClient.SqlCommand("SELECT Vouchers_Status from VouchersToSuppliers where Departure_Id =" & ss, myConnection)
                cmdSelect.CommandType = CommandType.Text
                myConnection.Open()

                myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    If Not IsDBNull(myReader("Vouchers_Status")) Then
                        If Trim(myReader("Vouchers_Status")) <> "מותאם" Then
                            result = ""
                            Exit While
                        Else
                            result = "מותאם סופית"
                        End If



                    End If

                End While
                'If result = "" Then
                '    result = "הכל"
                'End If
                myConnection.close()


                Return result
            End If
        End Function
        Function EmailToByImportance(ByVal ValueStr As String) As String
            Dim result As String
            Dim ss As String
            ss = ValueStr
            '  Response.Write("ss=" & ss)
            '  Response.End()
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            If ss <> "" Then

                Dim cmdSelect As New SqlClient.SqlCommand("SELECT EMAIL from Users  where  ImportanceId in(" & ss & ") and (Email<>'' or Email<>null)", myConnection)
                cmdSelect.CommandType = CommandType.Text
                myConnection.Open()

                myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    If Not IsDBNull(myReader("EMAIL")) Then
                        If result <> "" Then
                            result = result & ", " & myReader("EMAIL")
                        Else
                            result = myReader("EMAIL")
                        End If


                    End If

                End While
                'If result = "" Then
                '    result = "הכל"
                'End If
                myConnection.close()


                Return result
            End If
        End Function
        Function GetSelectSupplierUserName(ByVal ValueStr As String) As String
            Dim result As String
            Dim ss As String
            ss = ValueStr
            '  Response.Write("ss=" & ss)
            '  Response.End()
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            If ss <> "" Then

                Dim cmdSelect As New SqlClient.SqlCommand("SELECT supplier_Name1,supplier_Email1,supplier_Name2,supplier_Email2,supplier_Name3,supplier_Email3,supplier_Name4,supplier_Email4 from Suppliers  where  supplier_ID in(" & ss & ") order by supplier_Name", myConnection)
                cmdSelect.CommandType = CommandType.Text
                myConnection.Open()

                myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    If Not IsDBNull(myReader("supplier_Name1")) And Not IsDBNull(myReader("supplier_Email1")) Then
                        If result <> "" Then
                            result = result & ", " & myReader("supplier_Name1")
                        Else
                            result = myReader("supplier_Name1")
                        End If

                        If Not IsDBNull(myReader("supplier_Name2")) And Not IsDBNull(myReader("supplier_Email2")) Then
                            If Trim(myReader("supplier_Name2")) <> "" Then
                                If result <> "" Then
                                    result = result & ", " & myReader("supplier_Name2")
                                Else
                                    result = myReader("supplier_Name2")
                                End If
                            End If
                        End If


                        If Not IsDBNull(myReader("supplier_Name3")) And Not IsDBNull(myReader("supplier_Email3")) Then
                            If Trim(myReader("supplier_Name3")) <> "" Then
                                If result <> "" Then
                                    result = result & ", " & myReader("supplier_Name3")
                                Else
                                    result = myReader("supplier_Name3")
                                End If
                            End If

                        End If
                        If Not IsDBNull(myReader("supplier_Name4")) And Not IsDBNull(myReader("supplier_Email4")) Then
                            If Trim(myReader("supplier_Name4")) <> "" Then
                                If result <> "" Then
                                    result = result & ", " & myReader("supplier_Name4")
                                Else
                                    result = myReader("supplier_Name4")
                                End If
                            End If

                        End If

                    End If

                End While
                'If result = "" Then
                '    result = "הכל"
                'End If
                myConnection.close()


                Return result
            End If
        End Function
        Function GetSelectSupplierGUID(ByVal ValueStr As String) As String
            Dim result As String
            Dim ss As String
            ss = ValueStr
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            If ss <> "" Then
                Dim cmdSelect As New SqlClient.SqlCommand("SELECT GUID from Suppliers  where  supplier_ID in(" & ss & ") order by supplier_Name", myConnection)
                cmdSelect.CommandType = CommandType.Text
                myConnection.Open()
                myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    If Not IsDBNull(myReader("GUID")) Then
                        result = myReader("GUID")
                    End If
                End While
            End If
            myConnection.close()
            Return result
        End Function
        Function GetSuppliersByDepId(ByVal DepartureId As Integer) As String
            Dim result As String
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            'Dim cmdSelect As New SqlClient.SqlCommand("select VS.supplier_Id,S.supplier_Name,Departure_Id,Vouchers_Status ,Country_Name " & _
            '             " from VouchersToSuppliers VS   left join Suppliers S on VS.supplier_Id=S.supplier_Id " & _
            '             " left join pegasus.dbo.Countries on pegasus.dbo.Countries.Country_Id=S.Country_Id  where Departure_Id = " & CInt(DepartureId), myConnection)
            ''  Response.Write(cmdSelect.CommandText)
            Dim cmdSelect As New SqlClient.SqlCommand("select distinct VS.supplier_Id,S.supplier_Name,Departure_Id,Country_Name " & _
                                " from VouchersToSuppliers VS   left join Suppliers S on VS.supplier_Id=S.supplier_Id " & _
                                " left join " & PegasusSiteDBName & ".dbo.Countries on pegasus.dbo.Countries.Country_Id=S.Country_Id  where Departure_Id = " & CInt(DepartureId), myConnection)

            cmdSelect.CommandType = CommandType.Text
            myConnection.Open()
            myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
            While myReader.Read()
                If Not IsDBNull(myReader("supplier_Id")) Then
                    If result <> "" Then
                        result = result & ", " & myReader("supplier_Id")
                    Else
                        result = myReader("supplier_Id")
                    End If

                End If
            End While

            myConnection.Close()
            Return result

        End Function

        Function GetSelectSupplierUserEmail(ByVal ValueStr As String) As String
            Dim result As String
            Dim ss As String
            ss = ValueStr
            '  Response.Write("ss=" & ss)
            '  Response.End()
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            If ss <> "" Then

                Dim cmdSelect As New SqlClient.SqlCommand("SELECT supplier_Name1,supplier_Email1,supplier_Name2,supplier_Email2,supplier_Name3,supplier_Email3,supplier_Name4,supplier_Email4 from Suppliers  where  supplier_ID in(" & ss & ") order by supplier_Name", myConnection)
                cmdSelect.CommandType = CommandType.Text
                myConnection.Open()

                myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    If Not IsDBNull(myReader("supplier_Email1")) Then
                        If result <> "" Then
                            result = result & ", " & myReader("supplier_Email1")
                        Else
                            result = myReader("supplier_Email1")
                        End If
                    End If

                    If Not IsDBNull(myReader("supplier_Email2")) Then
                        If result <> "" Then
                            result = result & ", " & myReader("supplier_Email2")
                        Else
                            result = myReader("supplier_Email2")
                        End If
                    End If


                    If Not IsDBNull(myReader("supplier_Email3")) Then
                        If result <> "" Then
                            result = result & ", " & myReader("supplier_Email3")
                        Else
                            result = myReader("supplier_Email3")
                        End If
                    End If

                    If Not IsDBNull(myReader("supplier_Email4")) Then
                        If result <> "" Then
                            result = result & ", " & myReader("supplier_Email4")
                        Else
                            result = myReader("supplier_Email4")
                        End If
                    End If



                End While
                'If result = "" Then
                '    result = "הכל"
                'End If
                myConnection.close()


                Return result
            End If
        End Function



        Function GetSelectSupplierName(ByVal ValueStr As String) As String
            Dim result As String
            Dim ss As String
            ss = ValueStr
            '  Response.Write("ss=" & ss)
            '  Response.End()
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            If ss <> "" Then

                Dim cmdSelect As New SqlClient.SqlCommand("SELECT supplier_Name from Suppliers  where  supplier_ID in(" & ss & ") order by supplier_Name", myConnection)
                cmdSelect.CommandType = CommandType.Text
                myConnection.Open()

                myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    If Not IsDBNull(myReader("supplier_Name")) Then
                        If result <> "" Then
                            result = result & ", " & myReader("supplier_Name")
                        Else
                            result = myReader("supplier_Name")
                        End If


                    End If

                End While
                'If result = "" Then
                '    result = "הכל"
                'End If
                myConnection.close()


                Return result
            End If
        End Function
        Function GetSelectGuidesName(ByVal ValueStr As String) As String
            Dim result As String
            Dim ss As String
            ss = ValueStr
            '  Response.Write("ss=" & ss)
            '  Response.End()
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))

            If ss <> "" Then
                Dim cmdSelect As New SqlClient.SqlCommand("SELECT Guide_ID, Guide_FName + ' ' + Guide_LName as GuideName from Guides  where  Guide_ID in(" & ss & ") order by Guide_FName asc ,Guide_LName asc", myConnection)
                cmdSelect.CommandType = CommandType.Text
                myConnection.Open()

                myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    If Not IsDBNull(myReader("GuideName")) Then
                        If result <> "" Then
                            result = result & ", " & myReader("GuideName")
                        Else
                            result = myReader("GuideName")
                        End If


                    End If

                End While
                If result = "" Then
                    result = "הכל"
                End If
                myConnection.close()


                Return result
            End If
        End Function
        Function GetMainCountryNameByTour(ByVal ValueStr As String) As String
            Dim result As String
            Dim ss As String
            ss = ValueStr
            '  Response.Write("ss=" & ss)
            '  Response.End()
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))

            If ss <> "" Then

                Dim cmdSelect As New SqlClient.SqlCommand("Select top 1 Country_Name from  Countries left join  Tours_Countries  on Countries.Country_id=Tours_Countries.Country_id " & _
                "  where isMain=1 and Tours_Countries.Tour_Id in(" & ss & ") ", myConnection)
                cmdSelect.CommandType = CommandType.Text
                myConnection.Open()

                myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    If Not IsDBNull(myReader("Country_Name")) Then
                     
                            result = myReader("Country_Name")

                        End If

                    End While
                myConnection.close()
                Return result
            End If
        End Function
        Function GetSelectDepName(ByVal ValueStr As String) As String
            Dim result As String
            Dim ss As String
            ss = ValueStr
            '  Response.Write("ss=" & ss)
            '  Response.End()
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            If ss <> "" Then
                Dim cmdSelect As New SqlClient.SqlCommand("SELECT departmentName from Departments  where  departmentId in(" & ss & ") order by departmentName", myConnection)
                cmdSelect.CommandType = CommandType.Text
                myConnection.Open()

                myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    If Not IsDBNull(myReader("departmentName")) Then
                        If result <> "" Then
                            result = result & ", " & myReader("departmentName")
                        Else
                            result = myReader("departmentName")
                        End If
                    End If

                End While
                If result = "" Then
                    result = "הכל"
                End If
                myConnection.close()


                Return result
            End If
        End Function
        Function GetDepId(ByVal ValueStr As Integer) As String
            Dim result As String
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            If IsNumeric(ValueStr) Then
                Dim cmdSelect As New SqlClient.SqlCommand("SELECT  Department_Id from Users  where  USER_ID= " & ValueStr, myConnection)
                cmdSelect.CommandType = CommandType.Text
                myConnection.Open()

                myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    If Not IsDBNull(myReader("Department_Id")) Then

                        result = myReader("Department_Id")
                    End If



                End While
                If result = "" Then
                    ''    result = "הכל"
                End If
                myConnection.close()


                Return result
            End If

        End Function
        Function GetUserEmail(ByVal ValueStr As Integer) As String
            Dim result As String
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            If IsNumeric(ValueStr) Then
                Dim cmdSelect As New SqlClient.SqlCommand("SELECT  EMAIL from Users  where  USER_ID= " & ValueStr, myConnection)
                cmdSelect.CommandType = CommandType.Text
                myConnection.Open()

                myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    If Not IsDBNull(myReader("EMAIL")) Then

                        result = myReader("EMAIL")
                    End If



                End While
                If result = "" Then
                    ''    result = "הכל"
                End If
                myConnection.close()


                Return result
            End If

        End Function
        Function GetUserNameEng(ByVal ValueStr As Integer) As String
            Dim result As String
            Dim ss As String
            ss = ValueStr
            '  Response.Write("ss=" & ss)
            '  Response.End()
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            If ss <> "" Then
                Dim cmdSelect As New SqlClient.SqlCommand("SELECT USER_ID, FIRSTNAME_ENG + ' ' + LASTNAME_ENG as Username from Users  where  USER_ID= " & ss, myConnection)
                cmdSelect.CommandType = CommandType.Text
                myConnection.Open()

                myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    If Not IsDBNull(myReader("Username")) Then
                        If result <> "" Then
                            result = result & ", " & myReader("Username")
                        Else
                            result = myReader("Username")
                        End If


                    End If

                End While
                If result = "" Then
                    ''    result = "הכל"
                End If
                myConnection.close()


                Return result
            End If
        End Function
        Function GetSelectUserName(ByVal ValueStr As String) As String
            Dim result As String
            Dim ss As String
            ss = ValueStr
            '  Response.Write("ss=" & ss)
            '  Response.End()
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            If ss <> "" Then
                Dim cmdSelect As New SqlClient.SqlCommand("SELECT USER_ID, FIRSTNAME + ' ' + LASTNAME as Username from Users  where  USER_ID in(" & ss & ") order by FIRSTNAME asc ,LASTNAME asc", myConnection)
                cmdSelect.CommandType = CommandType.Text
                myConnection.Open()

                myReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)
                While myReader.Read()
                    If Not IsDBNull(myReader("Username")) Then
                        If result <> "" Then
                            result = result & ", " & myReader("Username")
                        Else
                            result = myReader("Username")
                        End If


                    End If

                End While
                If result = "" Then
                    result = "הכל"
                End If
                myConnection.close()


                Return result
            End If
        End Function
        Function QFix(ByVal myString As Object) As String
            Dim quote, newQuote, temp As String
            If Not IsDBNull(myString) Then
                temp = CStr(myString)
            Else
                Return ""
                Exit Function
            End If

            If temp <> "" Then
                quote = """"
                newQuote = "''"
                If InStr(myString, """") Then
                    temp = Replace(temp, quote, newQuote)
                End If
                temp = Replace(temp, """", "'+char(39)+'")
                temp = Replace(temp, Chr(150), Chr(45)) ' "-"
            End If
            Return temp

        End Function

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
        Function TourStatusForOperation(ByVal Departure_Id As String, ByVal Departure_DateBegin As String, ByVal Departure_DateEnd As String) As String
            Dim temp As String
            Dim today As DateTime = DateTime.Now
            Dim DateBegin As DateTime
            If IsDate(Departure_DateBegin) Then
                DateBegin = Convert.ToDateTime(Departure_DateBegin.ToString())
            End If

            Dim DateEnd As DateTime
            If IsDate(Departure_DateEnd) Then
                DateEnd = Convert.ToDateTime(Departure_DateEnd.ToString())
            End If

            If DateDiff(DateInterval.Day, today, DateBegin) >= 0 Then

                If DateDiff(DateInterval.Day, today, DateBegin) = 0 Then
                    temp = "יוצא היום "
                Else
                    '  temp = "יוצא בעוד " & DateDiff(DateInterval.Day, today, DateBegin) & " ימים "
                    temp = DateDiff(DateInterval.Day, today, DateBegin)

                End If

            ElseIf DateDiff(DateInterval.Day, today, DateBegin) < 0 Then

                If DateDiff(DateInterval.Day, today, DateEnd) < 0 Then
                    temp = "חזר"
                Else
                    temp = "כן"
                End If
            End If
            'temp = DateBegin & " :" & DateDiff(DateInterval.Day, today, DateBegin) & "<BR> " & DateEnd & ":" & DateDiff(DateInterval.Day, today, DateEnd)

            '  temp = Departure_DateBegin & ":" & Departure_DateEnd
            ' temp = DateDiff(DateInterval.Day, today, DateBegin)
            Return temp
        End Function
        Function GetSeriasId(ByVal User_Id As String) As String
            Dim StrSeriasId As String = ""

            Dim sqlstr = "Select Series_Id  from Series where User_Id=" & User_Id
            Dim myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            Dim myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()
            myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)

            While myReader.Read()
                If StrSeriasId = "" Then
                    StrSeriasId = myReader("Series_Id")
                Else
                    StrSeriasId = StrSeriasId & "," & myReader("Series_Id")
                End If
            End While
            myConnection.Close()

            Return StrSeriasId

        End Function


        Function GetSalesCount(ByVal dateM As String, ByVal questionsid As String, ByVal statusId As String) As String
            Dim sqlstr As String
            Dim tmp As Integer = 0
            'appeal_CallStatus= 1 'פניית שירות
            'else פניית מכירה

            If dateM <> "" Then
                'If questionsid = "16724" Then
                '    sqlstr = "SET DATEFORMAT dmy;Select IsNULL(COUNT(APPEAL_ID), 0) as s from Appeals where (questions_id=16724 or  questions_id=17012)  and appeal_CallStatus=" & statusId & "  and  datediff(d,APPEAL_Date,convert(smalldatetime, '" & CDate(dateM) & "',103))=0 "

                'Else
                sqlstr = "SET DATEFORMAT dmy;Select IsNULL(COUNT(APPEAL_ID), 0) as s from Appeals where questions_id=" & questionsid & " and appeal_CallStatus=" & statusId & "  and  datediff(d,APPEAL_Date,convert(smalldatetime, '" & CDate(dateM) & "',103))=0 "

                ' End If
                '  current.response.write(sqlstr)
                myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
                myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
                myConnection.Open()
                Dim Count = myCommand.ExecuteScalar
                If Count > 0 Then
                    tmp = Count
                Else
                    tmp = 0
                End If

                Return tmp

            End If

        End Function
        Function GetSalesCount16504(ByVal dateM As String, ByVal questionsid As String) As String
            Dim sqlstr As String
            Dim tmp As Integer = 0
            'appeal_CallStatus= 1 'פניית שירות
            'else פניית מכירה

            If dateM <> "" Then
                sqlstr = "SET DATEFORMAT dmy;Select IsNULL(COUNT(APPEAL_ID), 0) as s from Appeals where questions_id=" & questionsid & " and  Reason_Id=1 and  datediff(d,APPEAL_Date,convert(smalldatetime, '" & CDate(dateM) & "',103))=0 "
                'Current.response.write(sqlstr)
                'response.end()
                myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
                myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
                myConnection.Open()
                Dim Count = myCommand.ExecuteScalar
                If Count > 0 Then
                    tmp = Count
                Else
                    tmp = 0
                End If

                Return tmp

            End If

        End Function

        Public Function TimeZmanSales(ByVal currMonth As String, ByVal currYear As String, ByVal DepId As String)
            Dim sqlstr As String
            sqlstr = "SELECT TimeZmanSales from Departments_Data_Monthly where departmentId = " & DepId & " and  Dim_Month=" & currMonth & " and DimYear=" & currYear
            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()
            Dim value = myCommand.ExecuteScalar
            If IsNumeric(value) Then
            Else
                value = ""
            End If
            myConnection.Close()
            Return value

        End Function

        Public Function ProcSalesCallMerkaz(ByVal currMonth As String, ByVal currYear As String, ByVal DepId As String)
            Dim sqlstr As String
            sqlstr = "SELECT ProcSalesCallMerkaz from Departments_Data_Monthly where departmentId = " & DepId & " and  Dim_Month=" & currMonth & " and DimYear=" & currYear
            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()
            Dim value = myCommand.ExecuteScalar
            If IsNumeric(value) Then
            Else
                value = ""
            End If
            myConnection.Close()
            Return value

        End Function
        Public Function PotencialTo16504Sales(ByVal currMonth As String, ByVal currYear As String, ByVal DepId As String)
            'אחוז טפסי המתעניין ביחס לפוטנציאל המכירות (הצגת אחוז) table sales
            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            Dim sqlstr As String
            Dim sumPotencialSales, P16504 As Decimal
            Dim result As Decimal

            sqlstr = "SET DATEFORMAT DMY;SELECT IsNull(sum(P16719Kishurit_Sales),0) as P16719Kishurit_Sales," & _
                   " isNull(sum(P16724ContactUs_Sales),0) as P16724ContactUs_Sales,IsNull(sum(P17012ContactUs_Sales),0) as  P17012ContactUs_Sales," & _
                   " IsNull(sum(P16504),0) as P16504, IsNull(sum(DD.call_Sales),0) as call_Sales  " & _
                   " FROM DimDate A  left join Departments_Data DD on DD.DateKey=A.DateKey " & _
                   " WHERE CalendarYear = @currYear and CalendarMonth=@month and DD.departmentId=@dep"


            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myCommand.Parameters.Add("@currYear", SqlDbType.Int).Value = CInt(currYear)
            myCommand.Parameters.Add("@month", SqlDbType.Int).Value = CInt(currMonth)
            myCommand.Parameters.Add("@dep", SqlDbType.Int).Value = CInt(DepId)

            myConnection.Open()
            myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)
            While myReader.Read()
                If Not IsDBNull(myReader("P16719Kishurit_Sales")) Then
                    sumPotencialSales = sumPotencialSales + myReader("P16719Kishurit_Sales")
                End If
                If Not IsDBNull(myReader("P16724ContactUs_Sales")) Then
                    sumPotencialSales = sumPotencialSales + myReader("P16724ContactUs_Sales")
                End If

                If Not IsDBNull(myReader("P17012ContactUs_Sales")) Then
                    sumPotencialSales = sumPotencialSales + myReader("P17012ContactUs_Sales")
                End If
                If Not IsDBNull(myReader("call_Sales")) Then
                    sumPotencialSales = sumPotencialSales + myReader("call_Sales")
                End If
                If Not IsDBNull(myReader("P16504")) Then
                    P16504 = myReader("P16504")
                Else
                    P16504 = 0
                End If
                If sumPotencialSales > 0 Then
                    result = (myReader("P16504") / sumPotencialSales) * 100
                    'pP16504 = myReader("P16504")
                End If

            End While
            '
            Return result


        End Function
        Public Function AvrgService(ByVal currMonth As String, ByVal currYear As String, ByVal DepId As String)
            'היחס בין כמות קישוריות השירות לעומת כמות שיחות השירות שנכנסה (יציג אחוז) table service 
            Dim sqlstr, query As String
            Dim pMonthDays As Integer
            Dim result As Decimal
            pMonthDays = DateTime.DaysInMonth(currYear, currMonth)
            If DateDiff("d", "01/" & currMonth & "/" & currYear, Now()) < pMonthDays Then
                query = " and (Month(DimDate.DateKey) = " & currMonth & " And Year(DimDate.DateKey) = " & currYear & " And day(DimDate.DateKey) < " & Day(Now()) & ")"

            Else
                query = " and (Month(DimDate.DateKey) = " & currMonth & " And Year(DimDate.DateKey) = " & currYear & ")"

            End If
            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            sqlstr = "SET DATEFORMAT dmy;SELECT sum(Departments_Data.call_Service) as Scall_Service,sum(convert(decimal(6,2),DayTypeId)) as SDayTypeId " & _
     "  from DimDate  left join Departments_Data  on  DimDate.DateKey=Departments_Data.DateKey where departmentId = " & DepId & query
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()
            myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)


            While myReader.Read()
                If Not IsDBNull(myReader("Scall_Service")) And Not IsDBNull(myReader("SDayTypeId")) And myReader("SDayTypeId") > 0 Then

                    result = Math.Round(myReader("Scall_Service") / myReader("SDayTypeId"), 2)
                Else
                    result = 0
                End If

            End While
            myConnection.Close()

            If result > 0 Then
                Dim result_k As Decimal
                sqlstr = "SET DATEFORMAT dmy;SELECT sum(P16719Kishurit_Service) as P16719Kishurit_Service,sum(convert(decimal(6,2),DayTypeId)) as SDayTypeId " & _
                  "  from DimDate where 1=1" & query
                myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
                myConnection.Open()
                myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)


                While myReader.Read()
                    If Not IsDBNull(myReader("P16719Kishurit_Service")) Then

                        result_k = Math.Round(myReader("P16719Kishurit_Service") / myReader("SDayTypeId"), 2)
                    Else
                        result_k = 0
                    End If

                End While
                myConnection.Close()
                result = (result_k / result) * 100

            End If
            Return result

        End Function
        Public Function AvrgCallService(ByVal currMonth As String, ByVal currYear As String, ByVal DepId As String)
            Dim sqlstr, query As String
            Dim pMonthDays As Integer
            Dim result As Decimal
            pMonthDays = DateTime.DaysInMonth(currYear, currMonth)
            If DateDiff("d", "01/" & currMonth & "/" & currYear, Now()) < pMonthDays Then
                query = " and (Month(DimDate.DateKey) = " & currMonth & " And Year(DimDate.DateKey) = " & currYear & " And day(DimDate.DateKey) < " & Day(Now()) & ")"

            Else
                query = " and (Month(DimDate.DateKey) = " & currMonth & " And Year(DimDate.DateKey) = " & currYear & ")"

            End If


            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            sqlstr = "SET DATEFORMAT dmy;SELECT sum(Departments_Data.call_Service) as Scall_Service,sum(convert(decimal(6,2),DayTypeId)) as SDayTypeId " & _
     "  from DimDate  left join Departments_Data  on  DimDate.DateKey=Departments_Data.DateKey where departmentId = " & DepId & query
            '    response.write(sqlstr)
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()
            myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)


            While myReader.Read()
                If Not IsDBNull(myReader("Scall_Service")) And Not IsDBNull(myReader("SDayTypeId")) Then
                    If fixNumeric(myReader("SDayTypeId")) Then
                        result = Math.Round(myReader("Scall_Service") / myReader("SDayTypeId"), 2)
                    Else

                        result = 0

                    End If
                Else
                    result = 0
                End If

            End While
            myConnection.Close()
            Return result

        End Function
        Public Function ProcCallService(ByVal Kishuriot16719Service As Integer, ByVal AvrgCallService As Integer)
            Dim result As Decimal
            If fixNumeric(AvrgCallService) > 0 Then
                result = Math.Round((Kishuriot16719Service / AvrgCallService) * 100, 2)
            Else
                result = 0

            End If
            Return result

        End Function
        Public Function AvrgCallSales(ByVal currMonth As String, ByVal currYear As String, ByVal DepId As String)
            Dim sqlstr, query As String
            Dim pMonthDays As Integer
            Dim result As Decimal
            pMonthDays = DateTime.DaysInMonth(currYear, currMonth)

            If DateDiff("d", "01/" & currMonth & "/" & currYear, Now()) < pMonthDays Then
                query = " and (Month(DimDate.DateKey) = " & currMonth & " And Year(DimDate.DateKey) = " & currYear & " And day(DimDate.DateKey) < " & Day(Now()) & ")"

            Else
                query = " and (Month(DimDate.DateKey) = " & currMonth & " And Year(DimDate.DateKey) = " & currYear & ")"

            End If


            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            sqlstr = "SET DATEFORMAT dmy;SELECT sum(Departments_Data.call_Sales) as Scall_Sales,sum(convert(decimal(6,2),DayTypeId)) as SDayTypeId " & _
     "  from DimDate  left join Departments_Data  on  DimDate.DateKey=Departments_Data.DateKey where departmentId = " & DepId & query
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()
            myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)


            While myReader.Read()
                If Not IsDBNull(myReader("Scall_Sales")) Then
                    If (Not IsDBNull(myReader("SDayTypeId")) And myReader("SDayTypeId") > 0) Then
                        result = Math.Round(myReader("Scall_Sales") / myReader("SDayTypeId"), 2)
                    Else
                        result = 0
                    End If

                Else
                    result = 0
                End If

            End While
            myConnection.Close()
            Return result

        End Function

        Public Function Sales_Point(ByVal currMonth As String, ByVal currYear As String, ByVal DepId As String)
            Dim sqlstr As String
            sqlstr = "SELECT Sales_Point from Departments_Data_Monthly where departmentId = " & DepId & " and  Dim_Month=" & currMonth & " and DimYear=" & currYear
            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()
            Dim value = myCommand.ExecuteScalar
            If IsNumeric(value) Then
            Else
                value = "0"
            End If
            myConnection.Close()
            Return value

        End Function
        Public Function Kishuriot16719Service(ByVal currMonth As String, ByVal currYear As String)

            Dim sqlstr, query As String
            Dim pMonthDays As Integer
            Dim result As Decimal
            pMonthDays = DateTime.DaysInMonth(currYear, currMonth)
            If DateDiff("d", "01/" & currMonth & "/" & currYear, Now()) < pMonthDays Then
                query = " and (Month(DimDate.DateKey) = " & currMonth & " And Year(DimDate.DateKey) = " & currYear & " And day(DimDate.DateKey) < " & Day(Now()) & ")"

            Else
                query = " and (Month(DimDate.DateKey) = " & currMonth & " And Year(DimDate.DateKey) = " & currYear & ")"

            End If


            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            sqlstr = "SET DATEFORMAT dmy;SELECT sum(P16719Kishurit_Service) as P16719Kishurit_Service,sum(convert(decimal(6,2),DayTypeId)) as SDayTypeId " & _
     "  from DimDate where 1=1" & query
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()
            myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)


            While myReader.Read()
                If Not IsDBNull(myReader("P16719Kishurit_Service")) And Not IsDBNull(myReader("SDayTypeId")) Then
                    If fixNumeric(myReader("SDayTypeId")) > 0 Then
                        result = Math.Round(myReader("P16719Kishurit_Service") / myReader("SDayTypeId"), 2)
                    Else
                        result = 0
                    End If
                Else
                    result = 0
                End If

            End While
            myConnection.Close()
            Return result

        End Function
        Public Function Neto16504(ByVal currMonth As String, ByVal currYear As String)

            Dim sqlstr, query As String
            Dim pMonthDays As Integer
            Dim result As Decimal
            pMonthDays = DateTime.DaysInMonth(currYear, currMonth)
            If DateDiff("d", "01/" & currMonth & "/" & currYear, Now()) < pMonthDays Then
                query = " and (Month(DimDate.DateKey) = " & currMonth & " And Year(DimDate.DateKey) = " & currYear & " And day(DimDate.DateKey) < " & Day(Now()) & ")"

            Else
                query = " and (Month(DimDate.DateKey) = " & currMonth & " And Year(DimDate.DateKey) = " & currYear & ")"

            End If


            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            sqlstr = "SET DATEFORMAT dmy;SELECT sum(P16504) as P16504,sum(convert(decimal(6,2),DayTypeId)) as SDayTypeId " & _
     "  from DimDate where 1=1" & query
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()
            myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)


            While myReader.Read()
                If Not IsDBNull(myReader("P16504")) Then
                    If CDbl(myReader("SDayTypeId")) > 0 Then
                        result = Math.Round(myReader("P16504") / myReader("SDayTypeId"), 2)
                    Else
                        result = 0
                    End If

                Else
                    result = 0
                End If

            End While
            myConnection.Close()
            Return result

        End Function

        Public Function Neto16735(ByVal currMonth As String, ByVal currYear As String, ByVal DepId As String)

            Dim sqlstr As String
            sqlstr = "SELECT Sum(ISNULL(P16735,0))-Sum(ISNULL(P16735_Bitulim,0)) as V  from Departments_Data where departmentId = " & DepId & " and  Month(DateKey)=" & currMonth & " and year(DateKey)=" & currYear
            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()
            Dim value = myCommand.ExecuteScalar
            If IsNumeric(value) Then
            Else
                value = ""
            End If
            myConnection.Close()
            Return value
        End Function
        Public Function GetSalesCount16735Bitulim(ByVal DepartId As String, ByVal DateSale As String) As String
            Dim sqlstr As String

            sqlstr = "SELECT  sum(cast(FIELD_VALUE as Int))  as s   FROM bizpower_pegasus.dbo.FORM_VALUE  WHERE(FIELD_ID = 40660 And IsNumeric(Field_Value) = 1)" & _
            " and APPEAL_ID in (select APPEAL_ID FROM   bizpower_pegasus.dbo.APPEALS where Department_Id= " & DepartId & _
            " and  datediff(d,APPEAL_Date_Bitul,convert(smalldatetime, '" & DateSale & "',103))=0)"

            '  response.write(sqlstr)
            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()

            Dim CountMembers = myCommand.ExecuteScalar

            If IsNumeric(CountMembers) Then
            Else
                CountMembers = ""
            End If
            myConnection.Close()
            Return CountMembers
        End Function


        Public Function GetCountSaleByDate(ByVal DepartId As String, ByVal FieldId As String, ByVal DateSale As String) As String
            Dim sqlstr As String

            If FieldId = "40661" Then
                sqlstr = "SELECT  count(*)  as s   FROM bizpower_pegasus.dbo.FORM_VALUE  WHERE(FIELD_ID = " & FieldId & ")  And  (FIELD_VALUE = 'on') " & _
" and APPEAL_ID in (select APPEAL_ID FROM   bizpower_pegasus.dbo.APPEALS where Department_Id=" & DepartId & " and  datediff(d,APPEAL_Date,convert(smalldatetime, '" & DateSale & "',103))=0 )"

            Else
                sqlstr = "SELECT  sum(cast(FIELD_VALUE as Int))  as s   FROM bizpower_pegasus.dbo.FORM_VALUE  WHERE(FIELD_ID = " & FieldId & ") And IsNumeric(Field_Value) = 1 " & _
     " and APPEAL_ID in (select APPEAL_ID FROM   bizpower_pegasus.dbo.APPEALS where Department_Id=" & DepartId & " and  datediff(d,APPEAL_Date,convert(smalldatetime, '" & DateSale & "',103))=0 )"

            End If

            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()

            Dim CountMembers = myCommand.ExecuteScalar

            If IsNumeric(CountMembers) Then
            Else
                CountMembers = 0
            End If
            myConnection.Close()
            Return CountMembers

        End Function
        Function GetCountGuideMessages(ByVal Departure_Id As String) As String
            Dim sqlstr As String
            Dim tmp As String

            If Departure_Id <> "" Then
                sqlstr = "Select Count(*) as s from GuideMessages where Departure_Id=" & Departure_Id
                myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
                myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
                myConnection.Open()
                Dim Count = myCommand.ExecuteScalar
                If Count > 0 Then
                    tmp = Count
                Else
                    tmp = ""
                End If

                Return tmp

            End If


        End Function
        Function TourGradeCss(ByVal Departure_Id As String, ByVal BG As String) As String

            Dim sqlstr As String
            Dim result As Decimal
            Dim tmp As String

            If Departure_Id <> "" Then
                sqlstr = "Select sum(TourGrade) as Sum ,Count(*) as s from FeedBack_Form where feedback_status=2 and Departure_Id=" & Departure_Id
                myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
                myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
                myConnection.Open()
                myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)


                While myReader.Read()
                    If Not IsDBNull(myReader("Sum")) Then

                        result = Math.Round(myReader("Sum") / myReader("s"), 2)
                    Else
                        result = 0
                    End If

                End While
                myConnection.Close()
                Select Case result
                    Case 91.0 To 100
                        tmp = "background-color: rgb(0, 128, 0)"
                    Case 80.0 To 90.99
                        tmp = "background-color: rgb(0, 255, 0)"
                    Case 70.0 To 79.99
                        tmp = "background-color: rgb(255, 255, 0)"

                    Case 1 To 69.99
                        '    tmp = "background-color: rgb(248, 154, 80)"
                        tmp = "background-color: rgb(255, 0, 0)"
                    Case Else
                        tmp = BG

                        '    tmp = "background-color: #D8D8D8;"

                End Select
                'If result > 0 Then
                '    tmp = result
                'Else
                '    tmp = result
                'End If

                Return tmp

            End If


        End Function

        Function TourGrade(ByVal Departure_Id As String) As String

            Dim sqlstr As String
            Dim result As Decimal
            Dim tmp As String

            If Departure_Id <> "" Then
                sqlstr = "Select sum(TourGrade) as Sum ,Count(*) as s from FeedBack_Form where Departure_Id=" & Departure_Id & " and FeedBack_Status=2"
                myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
                myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
                myConnection.Open()
                myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)


                While myReader.Read()
                    If Not IsDBNull(myReader("Sum")) Then

                        result = Math.Round(myReader("Sum") / myReader("s"), 2)
                    Else
                        result = 0
                    End If

                End While
                myConnection.Close()
                If result > 0 Then
                    tmp = result & "%"
                Else
                    tmp = result
                End If

                Return tmp

            End If


        End Function


        Public Function Feedback_GetNextMinGrade(ByVal DepartureId As Integer, ByVal ContactId As Integer) As String
            Dim sqlstr As String
            sqlstr = "select top 1 *  from  (SELECT  top 2 Category_Name +' '+ Category_Name2  as CName,FeedBack_FormValue.radioF,FaqIndex   FROM            FeedBack_FormValue LEFT OUTER JOIN FeedBack_Form ON FeedBack_Form.Id = FeedBack_FormValue.FrmId " & _
                        " left join Feedback_FAQ ON Feedback_FAQ.FAQ_ID=FeedBack_FormValue.FaqIndex left join Feedback_Category on Feedback_Category.Category_id=Feedback_FAQ.Category_Id " & _
                        "  WHERE(FeedBack_FormValue.FaqIndex <= 14) and Contact_Id=" & ContactId & " and Departure_Id=" & DepartureId & _
                        " ORDER BY FeedBack_FormValue.FaqIndex) as s	 ORDER BY s.FaqIndex desc"
            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()
            myReader = myCommand.ExecuteReader(CommandBehavior.CloseConnection)
            Dim result As String

            While myReader.Read()
                If Not IsDBNull(myReader("CName")) Then
                    result = myReader("CName")
                Else
                    result = ""
                End If

            End While
            myConnection.Close()
            Return result

        End Function

        Public Function GetLastDateSale(ByVal DepartureCode As String) As String

            Dim sqlstr As String
            sqlstr = "SELECT   top 1   APPEAL_DATE FROM           bizpower_pegasus.dbo.Appeals  left join  bizpower_pegasus.dbo.Form_Value on bizpower_pegasus.dbo.Appeals.APPEAL_ID=bizpower_pegasus.dbo.Form_Value.APPEAL_ID  " & _
            " WHERE      (FIELD_ID = 40623) AND (FIELD_VALUE = '" & DepartureCode & " ') order by APPEAL_DATE desc"

            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()

            Dim CountMembers = myCommand.ExecuteScalar

            If IsDate(CountMembers) Then
            Else
                CountMembers = ""
            End If
            myConnection.Close()
            Return CountMembers

        End Function
        Public Function GetCountLastMonth(ByVal DepartureCode As String) As String
            Dim sqlstr As String
            sqlstr = "SELECT  sum(cast(FIELD_VALUE as Int))  as s   FROM bizpower_pegasus.dbo.FORM_VALUE  WHERE(FIELD_ID = 40660) And IsNumeric(Field_Value) = 1 " & _
" and APPEAL_ID in(select APPEAL_ID FROM  bizpower_pegasus.dbo.FORM_VALUE WHERE        (FIELD_ID = 40623) AND (FIELD_VALUE ='" & DepartureCode & "') ) " & _
" and APPEAL_ID in(select APPEAL_ID FROM  bizpower_pegasus.dbo.FORM_VALUE WHERE        (FIELD_ID = 40661) AND (FIELD_VALUE = '')" & _
" and APPEAL_ID in (select bizpower_pegasus.dbo.APPEALS.APPEAL_ID FROM  bizpower_pegasus.dbo.FORM_VALUE left join bizpower_pegasus.dbo.APPEALS on bizpower_pegasus.dbo.FORM_VALUE.APPEAL_ID=bizpower_pegasus.dbo.APPEALS.APPEAL_ID where " & _
 " APPEAL_Date >= DATEADD(Month,-1, GETDATE()))) "

            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()

            Dim CountMembers = myCommand.ExecuteScalar

            If IsNumeric(CountMembers) Then
            Else
                CountMembers = ""
            End If
            myConnection.Close()

            Return CountMembers

        End Function
        Public Function GetCountLast2Week(ByVal DepartureCode As String) As String
            Dim sqlstr As String
            sqlstr = "SELECT  sum(cast(FIELD_VALUE as Int))  as s   FROM bizpower_pegasus.dbo.FORM_VALUE  WHERE(FIELD_ID = 40660) And IsNumeric(Field_Value) = 1 " & _
" and APPEAL_ID in(select APPEAL_ID FROM  bizpower_pegasus.dbo.FORM_VALUE WHERE        (FIELD_ID = 40623) AND (FIELD_VALUE ='" & DepartureCode & "') ) " & _
" and APPEAL_ID in(select APPEAL_ID FROM  bizpower_pegasus.dbo.FORM_VALUE WHERE        (FIELD_ID = 40661) AND (FIELD_VALUE = '')" & _
" and APPEAL_ID in (select bizpower_pegasus.dbo.APPEALS.APPEAL_ID FROM  bizpower_pegasus.dbo.FORM_VALUE left join bizpower_pegasus.dbo.APPEALS on bizpower_pegasus.dbo.FORM_VALUE.APPEAL_ID=bizpower_pegasus.dbo.APPEALS.APPEAL_ID where " & _
 " APPEAL_Date >= DATEADD(day,-14, GETDATE()))) "
            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()

            Dim CountMembers = myCommand.ExecuteScalar

            If IsNumeric(CountMembers) Then
            Else
                CountMembers = ""
            End If
            myConnection.Close()
            Return CountMembers

        End Function

        Public Function GetCountLastDateSale(ByVal DepartureCode As String) As String
            Dim sqlstr As String
            sqlstr = "SELECT  sum(cast(FIELD_VALUE as Int))  as s   FROM bizpower_pegasus.dbo.FORM_VALUE  WHERE(FIELD_ID = 40660) And IsNumeric(Field_Value) = 1 " & _
" and APPEAL_ID in(select APPEAL_ID FROM  bizpower_pegasus.dbo.FORM_VALUE WHERE        (FIELD_ID = 40623) AND (FIELD_VALUE ='" & DepartureCode & "') ) " & _
" and APPEAL_ID in(select APPEAL_ID FROM  bizpower_pegasus.dbo.FORM_VALUE WHERE        (FIELD_ID = 40661) AND (FIELD_VALUE = '')" & _
" and APPEAL_ID in (select bizpower_pegasus.dbo.APPEALS.APPEAL_ID FROM  bizpower_pegasus.dbo.FORM_VALUE left join bizpower_pegasus.dbo.APPEALS on bizpower_pegasus.dbo.FORM_VALUE.APPEAL_ID=bizpower_pegasus.dbo.APPEALS.APPEAL_ID where " & _
 " APPEAL_Date >= DATEADD(day,-7, GETDATE()))) "

            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()

            Dim CountMembers = myCommand.ExecuteScalar

            If IsNumeric(CountMembers) Then
            Else
                CountMembers = ""
            End If
            myConnection.Close()
            Return CountMembers

        End Function

        Public Function GetDepatureCode(ByVal DepartureId As String) As String
            Dim sqlstr As String
            sqlstr = "SELECT Departure_Code   FROM Tours_Departures where Departure_Id=" & DepartureId
            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()
            Dim DepatureCode = myCommand.ExecuteScalar
            myConnection.Close()
            Return DepatureCode
        End Function

        Public Function GetTourUrl(ByVal TourId As String) As String
            Dim sqlstr As String
            sqlstr = "SELECT rewrite_friendlyUrl  FROM Tours where Tour_Id=" & TourId
            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()
            Dim friendlyURL = myCommand.ExecuteScalar
            myConnection.Close()
            Return friendlyURL
        End Function
        Public Function GetContactByDepBitulim(ByVal DepartureId As String, ByVal ContactId As String) As String
            Dim sqlstr As String
            sqlstr = "select  Count(*) as c FROM bizpower_pegasus.dbo.FORM_VALUE FV left join Appeals A on  FV.Appeal_Id=A.Appeal_Id " & _
            "   WHERE(FIELD_ID = 40661)  And  (FIELD_VALUE = 'on') and Departure_Id=@DepartureId and Contact_Id=@ContactId"
            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myCommand.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
            myCommand.Parameters.Add("@ContactId", SqlDbType.Int).Value = CInt(ContactId)
            myConnection.Open()
            Dim CountMembers = myCommand.ExecuteScalar

            If IsNumeric(CountMembers) Then
            Else
                CountMembers = ""
            End If
            myConnection.Close()
            Return CountMembers

        End Function
        Public Function GetCountMembersBitulim(ByVal DepartureCode As String) As String
            Dim sqlstr As String
            sqlstr = "SELECT  sum(cast(FIELD_VALUE as Int))  as s   FROM bizpower_pegasus.dbo.FORM_VALUE    WHERE(FIELD_ID = 40660)  And IsNumeric(Field_Value) = 1 " & _
            "  and APPEAL_ID in(select APPEAL_ID FROM  bizpower_pegasus.dbo.FORM_VALUE WHERE        (FIELD_ID = 40623) AND (FIELD_VALUE = '" & DepartureCode & "') ) " & _
            "  and APPEAL_ID in(select APPEAL_ID FROM  bizpower_pegasus.dbo.FORM_VALUE WHERE        (FIELD_ID = 40661) AND (FIELD_VALUE = 'on') )"

            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()

            Dim CountMembers = myCommand.ExecuteScalar

            If IsNumeric(CountMembers) Then
            Else
                CountMembers = ""
            End If
            myConnection.Close()
            Return CountMembers

        End Function
        Public Function GetCountMembers(ByVal DepartureCode As String) As String
            Dim sqlstr As String
            sqlstr = "SELECT  sum(cast(FIELD_VALUE as Int))  as s   FROM bizpower_pegasus.dbo.FORM_VALUE    WHERE(FIELD_ID = 40660)  And IsNumeric(Field_Value) = 1 " & _
            "  and APPEAL_ID in(select APPEAL_ID FROM  bizpower_pegasus.dbo.FORM_VALUE WHERE        (FIELD_ID = 40623) AND (FIELD_VALUE = '" & DepartureCode & "') ) " & _
            "  and APPEAL_ID in(select APPEAL_ID FROM  bizpower_pegasus.dbo.FORM_VALUE WHERE        (FIELD_ID = 40661) AND (FIELD_VALUE = '') )"

            myConnection = New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
            myCommand = New System.Data.SqlClient.SqlCommand(sqlstr, myConnection)
            myConnection.Open()

            Dim CountMembers = myCommand.ExecuteScalar

            If IsNumeric(CountMembers) Then
            Else
                CountMembers = ""
            End If
            myConnection.Close()
            Return CountMembers

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
        Function breaks(ByVal word)
            If Not IsDBNull(word) Then
                word = Replace(word, Chr(13) & Chr(10), "<br>")
                word = Replace(word, Chr(13), "<br>")
                word = Replace(word, Chr(10), "<br>")
                breaks = word
            Else

                breaks = ""
            End If
        End Function
        Function altFix(ByVal word As Object) As String
            Dim temp As String
            If Not IsDBNull(word) Then
                temp = CStr(word)
            Else
                Return ""
                Exit Function
            End If

            temp = Replace(temp, "'", "\x27")    ' JScript encode apostrophes 
            temp = Replace(temp, """", "\x22")   ' JScript encode double-quotes 
            temp = HttpUtility.HtmlEncode(temp)  ' encode chars special to HTML 

            Return temp

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
        Function StatusOperation(ByVal objVal As Object) As String
            If IsDBNull(objVal) Then
                StatusOperation = ""
            Else
                If CInt(objVal) <= 0 Then
                    StatusOperation = "כן"
                Else
                    StatusOperation = ""

                End If
            End If
        End Function

        Function dbNullFix(ByVal objVal As Object) As String
            If IsDBNull(objVal) Then
                dbNullFix = ""
            Else
                dbNullFix = Trim(objVal)
            End If
        End Function
        '  Function dbNullFixDecimal(ByVal objVal As Object) As String
        '     If IsDBNull(objVal) Then
        '        dbNullFixDecimal = 0
        '   Else
        '       dbNullFixDecimal = Trim(objVal)
        '    End If
        'End Function

        'Public Function fixNumeric(ByVal objVal As Object) As Double
        '    If IsNumeric(objVal) Then
        '        Return CDbl(objVal)
        '    Else
        '        Return 0
        '    End If
        'End Function
        Public Function fixNumeric(ByVal objVal As Object) As Double
            Dim num As Double
            Try

                If IsNothing(objVal) Then
                    num = 0
                Else
                    If IsNumeric(objVal) Then
                        num = CDbl(objVal)
                    Else
                        num = 0
                    End If
                End If
            Catch ex As Exception
                num = 0
            End Try
            fixNumeric = num
        End Function


    End Class
End Namespace


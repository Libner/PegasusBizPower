Imports System.Drawing
Imports WebChart
Imports System.Web.UI


Public Class reportYearGraph
    Inherits System.Web.UI.Page
    Protected ChartControl1 As WebChart.AreaChart

    Dim func As New bizpower.cfunc
    Protected dateStart, dateEnd, countryid, sqlQuery, title, pcountryid, countryName, label2, label1, label, appealsCount16504 As String
    Protected appealsCount16504_1, appealsCount16504_2 As String
    Protected quest_id, pquest_id, s16504, s16735, CountAll, CountAll1, CountAll2, pSum, pSum1, pSum2 As Integer



#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim con1 As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        countryid = Trim(Request("country_id"))
        quest_id = 16735
        pquest_id = 16504
        dateStart = DateValue(Trim(Request("dateStart")))
        dateEnd = DateValue(Trim(Request("dateEnd")))
        title = "השוואה של טפסי מתעניין ואחוזי סגירה לפי 3 שנים\n"
        title = title & "  תקופת  הייחוס " & dateStart & " - " & dateEnd

        sqlQuery = " and appeal_CountryId   in (" & countryid & ")"
        'response.Write "countryid-="&countryid   

        Dim sql = "Exec dbo.get_reportRegistrationYearsCountry  @start_date='" & dateStart & "', @end_date='" & dateEnd & "'," & _
          "@start_date1='" & DateAdd("yyyy", -1, dateStart) & "', @start_date2='" & DateAdd("yyyy", -2, dateStart) & "'," & _
          "@end_date1='" & DateAdd("yyyy", -1, dateEnd) & "', @end_date2='" & DateAdd("yyyy", -2, dateEnd) & "'," & _
          " @sqlQuery='" & sqlQuery & "', @productID=" & func.sFix(quest_id)
        'Response.Write(sql)
        'Response.End()

        Dim cmdSelect = New System.Data.SqlClient.SqlCommand(sql, con)
        con.Open()
        Dim drNew = cmdSelect.ExecuteReader()
        Do While drNew.Read()
            pcountryid = drNew("appeal_CountryId")
            countryName = func.QFix(drNew("Country_CRMName"))
    
            Dim sqlL = "Exec dbo.get_AppealsCount16504  @start_date='" & dateStart & "', @end_date='" & dateEnd & "'," & _
            " @countryID=" & func.sFix(pcountryid) & ", @productID=" & func.sFix(pquest_id)
            con1.Open()
            Dim cmdSelect1 = New System.Data.SqlClient.SqlCommand(sqlL, con1)
            Dim rs_count = cmdSelect1.ExecuteReader()
            If rs_count.Read() Then

                appealsCount16504 = CLng(rs_count(0))
            Else
                appealsCount16504 = ""
            End If
            rs_count = Nothing
            con1.Close()
            sqlL = "Exec dbo.get_AppealsCount16504  @start_date='" & DateAdd("yyyy", -1, dateStart) & "', @end_date='" & DateAdd("yyyy", -1, dateEnd) & "'," & _
                      " @countryID=" & func.sFix(pcountryid) & ", @productID=" & func.sFix(pquest_id)
            con1.Open()
            cmdSelect1 = New System.Data.SqlClient.SqlCommand(sqlL, con1)
            rs_count = cmdSelect1.ExecuteReader()
            If rs_count.Read() Then

                appealsCount16504_1 = CLng(rs_count(0))
            Else
                appealsCount16504_1 = ""
            End If
            rs_count = Nothing
            con1.Close()

            sqlL = "Exec dbo.get_AppealsCount16504  @start_date='" & DateAdd("yyyy", -2, dateStart) & "', @end_date='" & DateAdd("yyyy", -2, dateEnd) & "'," & _
                               " @countryID=" & func.sFix(pcountryid) & ", @productID=" & func.sFix(pquest_id)
            con1.Open()
            cmdSelect1 = New System.Data.SqlClient.SqlCommand(sqlL, con1)
            rs_count = cmdSelect1.ExecuteReader()
            If rs_count.Read() Then

                appealsCount16504_2 = CLng(rs_count(0))
            Else
                appealsCount16504_2 = ""
            End If
            rs_count = Nothing
            con1.Close()
            '------------------------
            '           sqlL = "SELECT distinct dbo.GetRegistrationCountAll ('16735',appeal_CountryId,'" & dateStart & "','" & dateEnd & "') as CountAll," & _
            '" dbo.GetRegistrationCountAll ('16735',appeal_CountryId,'" & DateAdd("yyyy", -1, dateStart) & "','" & DateAdd("yyyy", -1, dateEnd) & "' ) as CountAll_1," & _
            '" dbo.GetRegistrationCountAll ('16735',appeal_CountryId,'" & DateAdd("yyyy", -2, dateStart) & "','" & DateAdd("yyyy", -2, dateEnd) & "' ) as CountAll_2" & _
            '" FROM dbo.APPEALS APP " & _
            '" where 	 " & _
            '" APP.questions_id =" & func.sFix(quest_id) & " and appeal_CountryId=" & func.sFix(pcountryid)


            sqlL = "SELECT COUNT(APPEAL_ID) as  CountAll FROM APPEALS   where(appeal_CountryId = " & func.sFix(pcountryid) & " And QUESTIONS_ID = 16735 And RelationId > 0)" & _
" and (DateDiff(d,appeal_date, convert(smalldatetime,'" & dateStart & "' ,103)) <= 0 )	 " & _
"	AND (DateDiff(d,appeal_date, convert(smalldatetime,'" & dateEnd & "',103)) >= 0 )"


            '         Response.Write(sqlL)
            '        Response.End()
            con1.Open()
            cmdSelect1 = New System.Data.SqlClient.SqlCommand(sqlL, con1)
            rs_count = cmdSelect1.ExecuteReader()
            If rs_count.Read() Then
                CountAll = rs_count("CountAll")
            End If
            con1.Close()
            rs_count = Nothing


            '''    CountAll1 = rs_count("CountAll_1")
            '''    CountAll2 = rs_count("CountAll_2")
            '''End If
            pSum = Math.Round(CInt(CountAll) * 100 / CInt(appealsCount16504))

            sqlL = "SELECT COUNT(APPEAL_ID) as  CountAll1 FROM APPEALS   where(appeal_CountryId = " & func.sFix(pcountryid) & " And QUESTIONS_ID = 16735 And RelationId > 0)" & _
         " and (DateDiff(d,appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -1, dateStart) & "' ,103)) <= 0 )	 " & _
         "	AND (DateDiff(d,appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -1, dateEnd) & "',103)) >= 0 )"


            '         Response.Write(sqlL)
            '        Response.End()
            con1.Open()
            cmdSelect1 = New System.Data.SqlClient.SqlCommand(sqlL, con1)
            rs_count = cmdSelect1.ExecuteReader()
            If rs_count.Read() Then
                CountAll1 = rs_count("CountAll1")
            End If
            con1.Close()
            rs_count = Nothing



            If appealsCount16504_1 <> 0 Then
                pSum1 = Math.Round(CLng(CountAll1) * 100 / CLng(appealsCount16504_1))
            Else
                pSum1 = Math.Round(CLng(CountAll1) * 100 / 1)
            End If
            sqlL = "SELECT COUNT(APPEAL_ID) as  CountAll2 FROM APPEALS   where(appeal_CountryId = " & func.sFix(pcountryid) & " And QUESTIONS_ID = 16735 And RelationId > 0)" & _
      " and (DateDiff(d,appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -2, dateStart) & "' ,103)) <= 0 )	 " & _
      "	AND (DateDiff(d,appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -2, dateEnd) & "',103)) >= 0 )"


            '         Response.Write(sqlL)
            '        Response.End()
            con1.Open()
            cmdSelect1 = New System.Data.SqlClient.SqlCommand(sqlL, con1)
            rs_count = cmdSelect1.ExecuteReader()
            If rs_count.Read() Then
                CountAll2 = rs_count("CountAll2")
            End If
            con1.Close()
            rs_count = Nothing


            If appealsCount16504_2 <> 0 Then
                pSum2 = Math.Round(CLng(CountAll2) * 100 / CLng(appealsCount16504_2))
            Else
                pSum2 = Math.Round(CLng(CountAll2) * 100 / 1)
            End If


            '''rs_count = Nothing
            '''con1.Close()
            '  pSum = 100
            'pSum1 = 150
            ' pSum2 = 50
            '-------------------------



            label2 = label2 & "{label: """ & countryName & """, y: " & appealsCount16504_2 & ",indexLabel: """ & appealsCount16504_2 & "/" & pSum2 & "%"",indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
            label1 = label1 & "{label: """ & countryName & """, y:" & appealsCount16504_1 & ",indexLabel: """ & appealsCount16504_1 & "/" & pSum1 & "%"",indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
            label = label & "{label: """ & countryName & """, y: " & appealsCount16504 & ",indexLabel: """ & appealsCount16504 & "/" & pSum & "%"",indexLabelOrientation: ""vertical"",  indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"


        Loop

        con.Close()
        drNew.Close()
        If Len(label) > 0 Then
            label = Left(label, Len(label) - 1)
        End If
        If Len(label1) > 0 Then
            label1 = Left(label1, Len(label1) - 1)
        End If
        If Len(label2) > 0 Then
            label2 = Left(label2, Len(label2) - 1)
        End If
    End Sub

End Class

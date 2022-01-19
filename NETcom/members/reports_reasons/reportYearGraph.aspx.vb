Imports System.Drawing
Imports WebChart
Imports System.Web.UI


Public Class reportYearGraphWithReasons
    Inherits System.Web.UI.Page
    Protected ChartControl1 As WebChart.AreaChart

    Dim func As New bizpower.cfunc
    Protected dateStart, dateEnd, countryid, sqlQuery, title, pcountryid, countryName, label2, label1, label, appealsCount16504 As String
    Protected appealsCount16504_1, appealsCount16504_2 As String
    Protected quest_id, pquest_id, s16504, s16735, CountAll, CountAll1, CountAll2, pSum, pSum1, pSum2 As Integer

    Dim reasons As String


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
        appealsCount16504_1 = 0
        appealsCount16504_2 = 0
        appealsCount16504 = 0
        countryid = Trim(Request("country_id"))
        quest_id = 16735
        pquest_id = 16504
        dateStart = DateValue(Trim(Request("dateStart")))
        dateEnd = DateValue(Trim(Request("dateEnd")))
        reasons = ""
        'only for tofes mitan'en

        'הגורם ליצירת הטופס
        '---------------P2932 - type of form- REASONS ---------------
        'select creation reason-------------------------------
        'added by Mila 22/10/2019-----------------------------	
        reasons = Request.Form("reasons")
        reasons = Replace(reasons, " ", "")

        '-----------------------------------------------------------


        title = "השוואה של טפסי מתעניין ואחוזי סגירה לפי 3 שנים\n"
        title = title & "  תקופת  הייחוס " & dateStart & " - " & dateEnd

        '---------------P2932 - type of form- REASONS ---------------
        If reasons <> "" Then
            Dim titleStr As String
            titleStr = titleStr & ",  הגורם ליצירת הטופס: "
            Dim i As Integer = 0
            Dim sqlstr As String = "SELECT  Reason_Title FROM  Appeals_CreationReasons where QUESTIONS_ID=16504 and Reason_Id in (" & reasons & ") order by Reason_Order"

            Dim cmdtitlte As New System.Data.SqlClient.SqlCommand(sqlstr, con)
            con.Open()
            Dim rs_Reason As SqlClient.SqlDataReader
            rs_Reason = cmdtitlte.ExecuteReader
            Do While rs_Reason.Read
                If i > 0 Then
                    titleStr = titleStr & ", "
                End If
                titleStr = titleStr & rs_Reason("Reason_Title")
                i = i + 1
            Loop
            rs_Reason.Close()
            con.Close()
            If titleStr <> "" Then
                title = title & " " & titleStr
            End If
        End If

        sqlQuery = " and appeal_CountryId   in (" & countryid & ")"
        'response.Write "countryid-="&countryid   

        Dim sqlCountry = "SELECT distinct appeal_CountryId,dbo.get_Country_CRMName(appeal_CountryId) Country_CRMName FROM dbo.APPEALS APP where appeal_CountryId   in (" & countryid & ") order by  dbo.get_Country_CRMName(appeal_CountryId),appeal_CountryId"
        '  Response.Write(sqlCountry)
        ' Response.End()
        Dim cmdCountry = New System.Data.SqlClient.SqlCommand(sqlCountry, con)
        con.Open()
        Dim drCountry = cmdCountry.ExecuteReader()
        Do While drCountry.Read()
            pcountryid = drCountry("appeal_CountryId")
            countryName = func.QFix(drCountry("Country_CRMName"))



            pquest_id = 16504
            Dim sql = "Exec dbo.get_AppealsCount16504_Reasons  @start_date='" & dateStart & "', @end_date='" & dateEnd & "'," & _
        " @countryID=" & func.sFix(pcountryid) & ", @productID=" & func.sFix(pquest_id) & ", @reasons='" & reasons & "'"
            'response.Write sql
            'response.end
            Dim cmdSelect = New System.Data.SqlClient.SqlCommand(sql, con1)
            con1.Open()
            Dim rs_count = cmdSelect.ExecuteReader()
            appealsCount16504 = ""
            While rs_count.Read()
                appealsCount16504 = CLng(rs_count(0))
            End While

            rs_count = Nothing
            con1.Close()


            Dim sqlL = "Exec dbo.get_AppealsCount16504_Reasons  @start_date='" & DateAdd("yyyy", -1, dateStart) & "', @end_date='" & DateAdd("yyyy", -1, dateEnd) & "'," & _
           " @countryID=" & func.sFix(pcountryid) & ", @productID=" & func.sFix(pquest_id) & ", @reasons='" & reasons & "'"
            Dim cmdSelectL = New System.Data.SqlClient.SqlCommand(sqlL, con1)
            con1.Open()

            Dim rs_countL = cmdSelectL.ExecuteReader()
            appealsCount16504_1 = ""
            While rs_countL.read()
                appealsCount16504_1 = CLng(rs_countL(0))
            End While

            rs_countL = Nothing
            con1.Close()

            Dim sql2 = "Exec dbo.get_AppealsCount16504_Reasons  @start_date='" & DateAdd("yyyy", -2, dateStart) & "', @end_date='" & DateAdd("yyyy", -2, dateEnd) & "'," & _
           " @countryID=" & func.sFix(pcountryid) & ", @productID=" & func.sFix(pquest_id) & ", @reasons='" & reasons & "'"
            Dim cmdSelect2 = New System.Data.SqlClient.SqlCommand(sql2, con1)
            con1.Open()
            Dim rs_count2 = cmdSelect2.ExecuteReader()
            appealsCount16504_2 = ""
            While rs_count2.read()
                appealsCount16504_2 = CLng(rs_count2(0))
            End While

            rs_count2 = Nothing
            con1.Close()

            If reasons = "" Then
                sqlL = "SELECT COUNT(APPEAL_ID) as  CountAll FROM APPEALS   where(appeal_CountryId = " & func.sFix(pcountryid) & " And QUESTIONS_ID = 16735 And RelationId > 0)" & _
             " and (DateDiff(d,appeal_date, convert(smalldatetime,'" & dateStart & "' ,103)) <= 0 )	 " & _
             "	AND (DateDiff(d,appeal_date, convert(smalldatetime,'" & dateEnd & "',103)) >= 0 )"
            Else
                sqlL = "SELECT COUNT(APP.APPEAL_ID) as CountAll FROM APPEALS APP inner join APPEALS app16504 on APP.RelationId=app16504.APPEAL_ID " & _
                                  "   where(APP.appeal_CountryId = " & func.sFix(pcountryid) & " And APP.QUESTIONS_ID = 16735 And APP.RelationId > 0)" & _
                           " and (DateDiff(d,APP.appeal_date, convert(smalldatetime,'" & dateStart & "' ,103)) <= 0 )	 " & _
                           "	AND (DateDiff(d,APP.appeal_date, convert(smalldatetime,'" & dateEnd & "',103)) >= 0 )	 " & _
                            " AND app16504.Reason_Id in (" & reasons & ") "

            End If

            '         Response.Write(sqlL)
            '        Response.End()
            con1.Open()
            Dim cmdSelect1 = New System.Data.SqlClient.SqlCommand(sqlL, con1)
            rs_count = cmdSelect1.ExecuteReader()
            If rs_count.Read() Then
                CountAll = rs_count("CountAll")
            End If
            con1.Close()
            rs_count = Nothing


            '''    CountAll1 = rs_count("CountAll_1")
            '''    CountAll2 = rs_count("CountAll_2")
            '''End If
            If appealsCount16504 <> 0 Then
                pSum = Math.Round(CInt(CountAll) * 100 / CInt(appealsCount16504))
            Else
                pSum = Math.Round(CLng(CountAll) * 100 / 1)
            End If
            If reasons = "" Then
                sqlL = "SELECT COUNT(APPEAL_ID) as  CountAll1 FROM APPEALS   where(appeal_CountryId = " & func.sFix(pcountryid) & " And QUESTIONS_ID = 16735 And RelationId > 0)" & _
             " and (DateDiff(d,appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -1, dateStart) & "' ,103)) <= 0 )	 " & _
             "	AND (DateDiff(d,appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -1, dateEnd) & "',103)) >= 0 )"
            Else
                sqlL = "SELECT COUNT(APP.APPEAL_ID) as CountAll1 FROM APPEALS APP inner join APPEALS app16504 on APP.RelationId=app16504.APPEAL_ID " & _
                    "  where(APP.appeal_CountryId = " & func.sFix(pcountryid) & " And APP.QUESTIONS_ID = 16735 And APP.RelationId > 0)" & _
            " and (DateDiff(d,APP.appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -1, dateStart) & "' ,103)) <= 0 )	 " & _
            "	AND (DateDiff(d,APP.appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -1, dateEnd) & "',103)) >= 0 )" & _
                " AND app16504.Reason_Id in (" & reasons & ") "
            End If

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
            If reasons = "" Then
                sqlL = "SELECT COUNT(APPEAL_ID) as  CountAll2 FROM APPEALS   where(appeal_CountryId = " & func.sFix(pcountryid) & " And QUESTIONS_ID = 16735 And RelationId > 0)" & _
          " and (DateDiff(d,appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -2, dateStart) & "' ,103)) <= 0 )	 " & _
          "	AND (DateDiff(d,appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -2, dateEnd) & "',103)) >= 0 )"
            Else
                sqlL = "SELECT COUNT(APP.APPEAL_ID) as  CountAll2  FROM APPEALS APP inner join APPEALS app16504 on APP.RelationId=app16504.APPEAL_ID " & _
                    "  where(APP.appeal_CountryId = " & func.sFix(pcountryid) & " And APP.QUESTIONS_ID = 16735 And APP.RelationId > 0)" & _
          " and (DateDiff(d,APP.appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -2, dateStart) & "' ,103)) <= 0 )	 " & _
          "	AND (DateDiff(d,APP.appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -2, dateEnd) & "',103)) >= 0 ) " & _
                " AND app16504.Reason_Id in (" & reasons & ") "
            End If


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



            '''''''''        Dim sql = "Exec dbo.get_reportRegistrationYearsCountry  @start_date='" & dateStart & "', @end_date='" & dateEnd & "'," & _
            '''''''''          "@start_date1='" & DateAdd("yyyy", -1, dateStart) & "', @start_date2='" & DateAdd("yyyy", -2, dateStart) & "'," & _
            '''''''''          "@end_date1='" & DateAdd("yyyy", -1, dateEnd) & "', @end_date2='" & DateAdd("yyyy", -2, dateEnd) & "'," & _
            '''''''''          " @sqlQuery='" & sqlQuery & "', @productID=" & func.sFix(quest_id)
            '''''''''        'Response.Write(sql)
            '''''''''        'Response.End()

            '''''''''        Dim cmdSelect = New System.Data.SqlClient.SqlCommand(sql, con)
            '''''''''        con.Open()
            '''''''''        Dim drNew = cmdSelect.ExecuteReader()
            '''''''''        Do While drNew.Read()


            '''''''''            Dim sqlL = "Exec dbo.get_AppealsCount16504  @start_date='" & dateStart & "', @end_date='" & dateEnd & "'," & _
            '''''''''            " @countryID=" & func.sFix(pcountryid) & ", @productID=" & func.sFix(pquest_id)
            '''''''''            con1.Open()
            '''''''''            Dim cmdSelect1 = New System.Data.SqlClient.SqlCommand(sqlL, con1)
            '''''''''            Dim rs_count = cmdSelect1.ExecuteReader()
            '''''''''            If rs_count.Read() Then

            '''''''''                appealsCount16504 = CLng(rs_count(0))
            '''''''''            Else
            '''''''''                appealsCount16504 = ""
            '''''''''            End If
            '''''''''            rs_count = Nothing
            '''''''''            con1.Close()
            '''''''''            sqlL = "Exec dbo.get_AppealsCount16504  @start_date='" & DateAdd("yyyy", -1, dateStart) & "', @end_date='" & DateAdd("yyyy", -1, dateEnd) & "'," & _
            '''''''''                      " @countryID=" & func.sFix(pcountryid) & ", @productID=" & func.sFix(pquest_id)
            '''''''''            con1.Open()
            '''''''''            cmdSelect1 = New System.Data.SqlClient.SqlCommand(sqlL, con1)
            '''''''''            rs_count = cmdSelect1.ExecuteReader()
            '''''''''            If rs_count.Read() Then

            '''''''''                appealsCount16504_1 = CLng(rs_count(0))
            '''''''''            Else
            '''''''''                appealsCount16504_1 = ""
            '''''''''            End If
            '''''''''            rs_count = Nothing
            '''''''''            con1.Close()

            '''''''''            sqlL = "Exec dbo.get_AppealsCount16504  @start_date='" & DateAdd("yyyy", -2, dateStart) & "', @end_date='" & DateAdd("yyyy", -2, dateEnd) & "'," & _
            '''''''''                               " @countryID=" & func.sFix(pcountryid) & ", @productID=" & func.sFix(pquest_id)
            '''''''''            con1.Open()
            '''''''''            cmdSelect1 = New System.Data.SqlClient.SqlCommand(sqlL, con1)
            '''''''''            rs_count = cmdSelect1.ExecuteReader()
            '''''''''            If rs_count.Read() Then

            '''''''''                appealsCount16504_2 = CLng(rs_count(0))
            '''''''''            Else
            '''''''''                appealsCount16504_2 = ""
            '''''''''            End If
            '''''''''            rs_count = Nothing
            '''''''''            con1.Close()
            '''''''''            '------------------------
            '''''''''    -----------        '           sqlL = "SELECT distinct dbo.GetRegistrationCountAll ('16735',appeal_CountryId,'" & dateStart & "','" & dateEnd & "') as CountAll," & _
            '''''''''            '" dbo.GetRegistrationCountAll ('16735',appeal_CountryId,'" & DateAdd("yyyy", -1, dateStart) & "','" & DateAdd("yyyy", -1, dateEnd) & "' ) as CountAll_1," & _
            '''''''''            '" dbo.GetRegistrationCountAll ('16735',appeal_CountryId,'" & DateAdd("yyyy", -2, dateStart) & "','" & DateAdd("yyyy", -2, dateEnd) & "' ) as CountAll_2" & _
            '''''''''            '" FROM dbo.APPEALS APP " & _
            '''''''''            '" where 	 " & _
            '''''''''            '" APP.questions_id =" & func.sFix(quest_id) & " and appeal_CountryId=" & func.sFix(pcountryid)


            '''''''''            sqlL = "SELECT COUNT(APPEAL_ID) as  CountAll FROM APPEALS   where(appeal_CountryId = " & func.sFix(pcountryid) & " And QUESTIONS_ID = 16735 And RelationId > 0)" & _
            '''''''''" and (DateDiff(d,appeal_date, convert(smalldatetime,'" & dateStart & "' ,103)) <= 0 )	 " & _
            '''''''''"	AND (DateDiff(d,appeal_date, convert(smalldatetime,'" & dateEnd & "',103)) >= 0 )"


            '''''''''            '         Response.Write(sqlL)
            '''''''''            '        Response.End()
            '''''''''            con1.Open()
            '''''''''            cmdSelect1 = New System.Data.SqlClient.SqlCommand(sqlL, con1)
            '''''''''            rs_count = cmdSelect1.ExecuteReader()
            '''''''''            If rs_count.Read() Then
            '''''''''                CountAll = rs_count("CountAll")
            '''''''''            End If
            '''''''''            con1.Close()
            '''''''''            rs_count = Nothing


            '''''''''            '''    CountAll1 = rs_count("CountAll_1")
            '''''''''            '''    CountAll2 = rs_count("CountAll_2")
            '''''''''            '''End If
            '''''''''            pSum = Math.Round(CInt(CountAll) * 100 / CInt(appealsCount16504))

            '''''''''            sqlL = "SELECT COUNT(APPEAL_ID) as  CountAll1 FROM APPEALS   where(appeal_CountryId = " & func.sFix(pcountryid) & " And QUESTIONS_ID = 16735 And RelationId > 0)" & _
            '''''''''         " and (DateDiff(d,appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -1, dateStart) & "' ,103)) <= 0 )	 " & _
            '''''''''         "	AND (DateDiff(d,appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -1, dateEnd) & "',103)) >= 0 )"


            '''''''''            '         Response.Write(sqlL)
            '''''''''            '        Response.End()
            '''''''''            con1.Open()
            '''''''''            cmdSelect1 = New System.Data.SqlClient.SqlCommand(sqlL, con1)
            '''''''''            rs_count = cmdSelect1.ExecuteReader()
            '''''''''            If rs_count.Read() Then
            '''''''''                CountAll1 = rs_count("CountAll1")
            '''''''''            End If
            '''''''''            con1.Close()
            '''''''''            rs_count = Nothing



            '''''''''            If appealsCount16504_1 <> 0 Then
            '''''''''                pSum1 = Math.Round(CLng(CountAll1) * 100 / CLng(appealsCount16504_1))
            '''''''''            Else
            '''''''''                pSum1 = Math.Round(CLng(CountAll1) * 100 / 1)
            '''''''''            End If
            '''''''''            sqlL = "SELECT COUNT(APPEAL_ID) as  CountAll2 FROM APPEALS   where(appeal_CountryId = " & func.sFix(pcountryid) & " And QUESTIONS_ID = 16735 And RelationId > 0)" & _
            '''''''''      " and (DateDiff(d,appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -2, dateStart) & "' ,103)) <= 0 )	 " & _
            '''''''''      "	AND (DateDiff(d,appeal_date, convert(smalldatetime,'" & DateAdd("yyyy", -2, dateEnd) & "',103)) >= 0 )"


            '''''''''            '         Response.Write(sqlL)
            '''''''''            '        Response.End()
            '''''''''            con1.Open()
            '''''''''            cmdSelect1 = New System.Data.SqlClient.SqlCommand(sqlL, con1)
            '''''''''            rs_count = cmdSelect1.ExecuteReader()
            '''''''''            If rs_count.Read() Then
            '''''''''                CountAll2 = rs_count("CountAll2")
            '''''''''            End If
            '''''''''            con1.Close()
            '''''''''            rs_count = Nothing


            '''''''''            If appealsCount16504_2 <> 0 Then
            '''''''''                pSum2 = Math.Round(CLng(CountAll2) * 100 / CLng(appealsCount16504_2))
            '''''''''            Else
            '''''''''                pSum2 = Math.Round(CLng(CountAll2) * 100 / 1)
            '''''''''            End If


            '''rs_count = Nothing
            '''con1.Close()
            'pSum = 100
            'pSum1 = 150
            'pSum2 = 50
            '-------------------------



            label2 = label2 & "{label: """ & countryName & """, y: " & appealsCount16504_2 & ",indexLabel: """ & appealsCount16504_2 & "/" & pSum2 & "%"",indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
            label1 = label1 & "{label: """ & countryName & """, y:" & appealsCount16504_1 & ",indexLabel: """ & appealsCount16504_1 & "/" & pSum1 & "%"",indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
            label = label & "{label: """ & countryName & """, y: " & appealsCount16504 & ",indexLabel: """ & appealsCount16504 & "/" & pSum & "%"",indexLabelOrientation: ""vertical"",  indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"


        Loop

        con.Close()
        ' drNew.Close()
        drCountry.Close()
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

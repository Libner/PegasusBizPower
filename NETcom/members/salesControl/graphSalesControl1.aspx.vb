Public Class graphSalesControl1
    Inherits System.Web.UI.Page
    Dim func As New include.funcs
    Protected currentYear, res1, res2, title, label2, label1, pMonthName As String
    Protected sumLabel1, sumLabel2, proc As Decimal


    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim rs_Guide, rs_r As System.Data.SqlClient.SqlDataReader

    Dim sqlstrGuide As New SqlClient.SqlCommand

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        currentYear = Request("currentYear")

        '  For i As Integer = 12 To 1 Step -1
        For i As Integer = 1 To 12
            ' MonthName = Month(i)
            pMonthName = MonthName(i)

            ' title = "דוח מכירות יעד מול ביצוע שנת " & currentYear
            con.Open()
            Dim sql As New SqlClient.SqlCommand("SELECT sum(IsNull(Sales,0)) as SalesYear, sum(isNull(Sales_Point,0)) as Sales_Point  " & _
            " from Departments_Data_Monthly WHERE DimYear = " & currentYear & " and Dim_Month=" & i, con)
            rs_r = sql.ExecuteReader(CommandBehavior.CloseConnection)
            'Response.Write(sql.CommandText & "<BR>")

            While rs_r.Read()
                If IsNumeric(rs_r("SalesYear")) Then
                    res1 = rs_r("SalesYear")
                Else
                    res1 = 0
                End If
                If IsNumeric(rs_r("Sales_Point")) Then
                    res2 = rs_r("Sales_Point")
                Else
                    res2 = 0
                End If
                'Response.Write("'" & res1 & ":" & res2 & "<BR>")
                'Response.End()
                If res2 > 0 Then
                    proc = (res1 / res2) * 100
                Else
                    proc = 0
                End If
                label2 = label2 & "{label: """ & pMonthName & " " & String.Format("{0:0.0}", proc) & "%" & """, y: " & res2 & ",indexLabel: """ & res2 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
                label1 = label1 & "{label: """ & pMonthName & " " & String.Format("{0:0.0}", proc) & "%" & """, y:" & res1 & ",indexLabel: """ & res1 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
                If IsNumeric(res1) Then
                    sumLabel1 = sumLabel1 + res1
                End If
                If IsNumeric(res2) Then
                    sumLabel2 = sumLabel2 + res2
                End If
            End While
            con.Close()
        Next
        ' Response.End()
        label2 = label2 & "{label: """ & "שנתי" & """, y: " & sumLabel2 & ",indexLabel: """ & sumLabel2 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
        label1 = label1 & "{label: """ & "שנתי" & """, y:" & sumLabel1 & ",indexLabel: """ & sumLabel1 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"


        'Dim s = "Exec dbo.[get_GuideFeedbackReport]  '" & guideId & "','" & FromDate & "','" & ToDate & "'"
        ''Response.Write(s)
        ''' Response.End()
        'Dim sql As New SqlClient.SqlCommand("Exec dbo.[get_GuideFeedbackReport]  '" & guideId & "','" & FromDate & "','" & ToDate & "'", conPegasus)
        'rs_r = sql.ExecuteReader(CommandBehavior.CloseConnection)
        'While rs_r.Read()
        '    res1 = rs_r("Tour_Grade")
        '    res2 = rs_r("Guide_Grade")
        '    '  label2 = label2 & "{label: """ & rs_r("Departure_Code") & " / מס משובים " & rs_r("CountFeedBack") & """, y: " & res2 & ",indexLabel: """ & res2 & "%"",indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
        '    ' label1 = label1 & "{label: """ & rs_r("Departure_Code") & """, y:" & res1 & ",indexLabel: """ & res1 & "%"",indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"


        '    label2 = label2 & "{label: """ & " מס משובים - " & rs_r("CountFeedBack") & """, y: " & res2 & ",indexLabel: """ & res2 & "%"",indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
        '    label1 = label1 & "{label: """ & rs_r("Departure_Code") & """, y:" & res1 & ",indexLabel:"" " & res1 & "% " & rs_r("Departure_Code") & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"


        '    ' label2 = label2 & "{label: """ & rs_r("Departure_Code") & """, y: " & res2 & ",indexLabel: """ & res2 & "%"",indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
        '    ' label1 = label1 & "{label: """ & rs_r("Departure_Code") & """, y:" & res1 & ",indexLabel: """ & res1 & "%"",indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"


        'End While


    End Sub
End Class

Public Class graphSalesControl3
    Inherits System.Web.UI.Page
    Dim func As New include.funcs
    Protected currentYear, res1, res2, title, label2, label1, pMonthName As String
    Protected sumLabel1, sumLabel2 As Decimal
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim rs_Guide, rs_r As System.Data.SqlClient.SqlDataReader

    Dim sqlstrGuide As New SqlClient.SqlCommand

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        currentYear = Request("currentYear")

        ' For i As Integer = 12 To 1 Step -1
        For i As Integer = 1 To 12
            ' MonthName = Month(i)
            pMonthName = MonthName(i)

            ' title = "דוח מכירות יעד מול ביצוע שנת " & currentYear
            con.Open()
            'Dim sql As New SqlClient.SqlCommand("SELECT sum(IsNull(Sales,0)) as SalesYear  " & _
            '" from Departments_Data_Monthly WHERE DimYear = " & currentYear & " and Dim_Month=" & i, conPegasus)
            Dim sql As New SqlClient.SqlCommand("SELECT DimYear,sum(IsNull(Sales,0)) as SalesYear " & _
                    " from Departments_Data_Monthly WHERE (DimYear = " & currentYear - 1 & " or DimYear=" & currentYear & ") and Dim_Month=" & i & "  group by DimYear order by DimYear", con)


            '        SELECT DimYear,Dim_Month,sum(IsNull(Sales,0)) as SalesYear 
            '      from  WHERE (DimYear=2016 or DimYear=currentYear) and Dim_Month=11
            'group by DimYear,Dim_Month

            rs_r = sql.ExecuteReader(CommandBehavior.CloseConnection)
            While rs_r.Read()
                If rs_r("DimYear") = currentYear Then
                    res1 = rs_r("SalesYear")
                Else
                    res2 = rs_r("SalesYear")
                End If

                ' res2 = rs_r("Sales_Point")


            End While
            If IsNumeric(res1) Then
                sumLabel1 = sumLabel1 + res1
            End If
            If IsNumeric(res2) Then
                sumLabel2 = sumLabel2 + res2
            End If
            label2 = label2 & "{label: """ & pMonthName & """, y: " & res2 & ",indexLabel: """ & res2 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
            label1 = label1 & "{label: """ & pMonthName & """, y:" & res1 & ",indexLabel: """ & res1 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"

            con.Close()
        Next

        label2 = label2 & "{label: """ & "שנתי" & """, y: " & sumLabel2 & ",indexLabel: """ & sumLabel2 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
        label1 = label1 & "{label: """ & "שנתי" & """, y:" & sumLabel1 & ",indexLabel: """ & sumLabel1 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"



    End Sub
End Class

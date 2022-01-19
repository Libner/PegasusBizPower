Public Class graphSalesControl4
    Inherits System.Web.UI.Page
    Dim func As New include.funcs
    Protected currentYear, res1, res2, title, label2, label1, pMonthName As String
    Protected sumLabel1, sumLabel2, proc, sumProc As Decimal
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim rs_Guide, rs_r As System.Data.SqlClient.SqlDataReader

    Dim sqlstrGuide As New SqlClient.SqlCommand

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        currentYear = Request("currentYear")
        '--טופס מתעניין בטיול--16504

        For i As Integer = 1 To 12
            ' MonthName = Month(i)
            pMonthName = MonthName(i)

            ' title = "דוח מכירות יעד מול ביצוע שנת " & currentYear
            con.Open()
            Dim sql As New SqlClient.SqlCommand("SET DATEFORMAT DMY;SELECT sum(IsNull(P16504,0) )as P16504, CalendarYear FROM DimDate  WHERE (CalendarYear =" & currentYear & _
            " or CalendarYear =" & currentYear - 1 & " ) " & _
            " and CalendarMonth=" & i & " group by CalendarYear order by CalendarYear", con)


            '        SELECT DimYear,Dim_Month,sum(IsNull(Sales,0)) as SalesYear 
            '      from Departments_Data_Monthly WHERE (DimYear=2016 or DimYear=currentYear) and Dim_Month=11
            'group by DimYear,Dim_Month

            rs_r = sql.ExecuteReader(CommandBehavior.CloseConnection)
            While rs_r.Read()
                If rs_r("CalendarYear") = currentYear Then
                    res1 = rs_r("P16504")
                Else
                    res2 = rs_r("P16504")
                End If

                ' res2 = rs_r("Sales_Point")


            End While
            If IsNumeric(res1) Then
                sumLabel1 = sumLabel1 + res1
            End If
            If IsNumeric(res2) Then
                sumLabel2 = sumLabel2 + res2
            End If
            If res2 > 0 Then
                proc = (res1 / res2) * 100
            Else
                proc = 0
            End If
            'label2 = label2 & "{label: """ & pMonthName & " " & String.Format("{0:0.0}", proc) & "%" & """, y: " & res2 & ",indexLabel: """ & res2 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
            'label1 = label1 & "{label: """ & pMonthName & " " & String.Format("{0:0.0}", proc) & "%" & """,, y:" & res1 & ",indexLabel: """ & res1 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"

            label2 = label2 & "{label: """ & pMonthName & " " & String.Format("{0:0.0}", proc) & "%" & """, y: " & res2 & ",indexLabel: """ & res2 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
            label1 = label1 & "{label: """ & pMonthName & " " & String.Format("{0:0.0}", proc) & "%" & """, y:" & res1 & ",indexLabel: """ & res1 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"



            con.Close()
        Next
        If sumLabel2 > 0 Then
            sumProc = (sumLabel1 / sumLabel2) * 100
        Else
            sumProc = 0
        End If

        '   label2 = label2 & "{label: """ & "שנתי" & " " & String.Format("{0:0.0}", sumProc) & "%" & """, y: " & sumLabel2 & ",indexLabel: """ & sumLabel2 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
        '   label1 = label1 & "{label: """ & "שנתי" & " " & String.Format("{0:0.0}", sumProc) & "%" & """,, y:" & sumLabel1 & ",indexLabel: """ & sumLabel1 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
        label2 = label2 & "{label: """ & "שנתי" & """, y: " & sumLabel2 & ",indexLabel: """ & sumLabel2 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"",indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
        label1 = label1 & "{label: """ & "שנתי" & """, y:" & sumLabel1 & ",indexLabel: """ & sumLabel1 & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"




    End Sub
End Class

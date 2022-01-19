Public Class graphSalesControl2
    Inherits System.Web.UI.Page
    Dim func As New include.funcs
    Protected currentYear, res1, res2, title, label2, label1, pQUARTERName As String
    Protected pRes As Integer
    Protected sumLabel1, sumLabel2, proc As Decimal


    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim rs_Guide, rs_r As System.Data.SqlClient.SqlDataReader

    Dim sqlstrGuide As New SqlClient.SqlCommand

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
        currentYear = Request("currentYear")
        ' Response.Write("currentYear=" & currentYear)
        ' Response.End()
        Dim pMonth As Integer
        Dim query As String
        'For i As Integer = 4 To 1 Step -1
        For i As Integer = 1 To 4
            ' MonthName = Month(i)
            pQUARTERName = " רבעון " & i
            Select Case i
                Case 1
                    query = " and (Dim_Month>=1 and Dim_Month<=3) "
                Case 2
                    query = " and (Dim_Month>=4 and Dim_Month<=6) "
                Case 3
                    query = " and (Dim_Month>=7 and Dim_Month<=9) "
                Case 4
                    query = " and (Dim_Month>=10 and Dim_Month<=12) "
            End Select
            ' title = "דוח מכירות יעד מול ביצוע שנת " & currentYear
            '  Dim s = "SELECT sum(IsNull(Sales,0)) as SalesYear, sum(isNull(Sales_Point,0)) as Sales_Point  " & _
            '  " from Departments_Data_Monthly WHERE DimYear = " & currentYear & " " & query
            '  Response.Write(s)
            '  Response.End()
            con.Open()
            Dim sql As New SqlClient.SqlCommand("SELECT sum(IsNull(Sales,0)) as SalesYear, sum(isNull(Sales_Point,0)) as Sales_Point  " & _
            " from Departments_Data_Monthly WHERE DimYear = " & currentYear & " " & query, con)
            rs_r = sql.ExecuteReader(CommandBehavior.CloseConnection)
            While rs_r.Read()
                res1 = rs_r("SalesYear")
                res2 = rs_r("Sales_Point")
                pRes = res1 - res2
                proc = (res1 / res2) * 100
                ' label1 = label1 & "{label: """ & pQUARTERName & """, y:" & pRes & ",indexLabel: """ & pRes & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"
                label1 = label1 & "{label: """ & pQUARTERName & """, y:" & pRes & ",indexLabel: """ & pRes & """,indexLabelOrientation: ""vertical"", indexLabelFontColor:""black"", indexLabelFontSize:""18"",indexlabelFontFamily: ""arial"",indexLabelPlacement: ""inside""},"

            End While
            con.Close()
        Next


    End Sub
End Class

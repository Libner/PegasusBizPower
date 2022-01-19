Public Class graphSalesControl6
    Inherits System.Web.UI.Page
    Dim func As New include.funcs
    Protected currentYear, res1, res2, title, label2, label1, pMonthName As String
    Protected sumLabel1, sumLabel2, proc, sumProc As Decimal
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim rs_Guide, rs_r As System.Data.SqlClient.SqlDataReader

    Dim sqlstrGuide As New SqlClient.SqlCommand


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
        currentYear = Request("currentYear")
        For i As Integer = 1 To 12
            ' MonthName = Month(i)
            pMonthName = MonthName(i)
            res1 = 0
            res2 = 0
            ' title = "��� ������ ��� ��� ����� ��� " & currentYear
            con.Open()
            Dim sql As New SqlClient.SqlCommand("SET DATEFORMAT DMY;SELECT sum(IsNull(call_Sales,0) )as call_Sales, Year(DateKey) as Year  FROM Departments_Data " & _
            " WHERE(Year(DateKey) = " & currentYear & " Or Year(DateKey) = " & currentYear - 1 & ")  and Month(DateKey)=" & i & " group by Year(DateKey)", con)
            rs_r = sql.ExecuteReader(CommandBehavior.CloseConnection)
            While rs_r.Read()
                If rs_r("Year") = currentYear Then
                    res1 = rs_r("call_Sales")
                Else
                    res2 = rs_r("call_Sales")
                End If
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


    End Sub

End Class

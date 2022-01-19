Public Class UpdateKishurit16719
    Inherits System.Web.UI.Page

    Public func As New bizpower.cfunc
    Public dr_m As SqlClient.SqlDataReader
    Protected dateM As String
    Protected P16719Kishurit_Sales, P16719Kishurit_Service, P16724ContactUs_Sales, P16724ContactUs_Service, P16504 As String
    Public pYear, pMonth As String



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
        pYear = Request.QueryString("y")
        pMonth = Request.QueryString("p")

        Dim sql As String

        'Put user code to initialize the page here
        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        If IsNumeric(pYear) Then
            sql = "SET DATEFORMAT dmy;SELECT  DateKey from DimDate  where year(DateKey)=" & pYear & " and month(DateKey)=" & pMonth
        Else
            sql = "SET DATEFORMAT dmy;SELECT  DateKey from DimDate  where (DateDiff(dd, DateKey, GetDate()) = 1)"
        End If
        Dim cmdSelect As New SqlClient.SqlCommand(sql, con)

        cmdSelect.CommandType = CommandType.Text
        con.Open()
        dr_m = cmdSelect.ExecuteReader()
        Dim urlF
        While dr_m.Read
            dateM = dr_m("DateKey")
            ' Response.Write("d=" & dateM & "<bR>")
            ' Response.End()
            'appeal_CallStatus= 1 'פניית שירות
            'else פניית מכירה
            'Dim questionsid, statusId



            '     questionsid = 16719
            '  statusId = 1

            Dim sqls As String

            ' Try
            P16719Kishurit_Sales = func.GetSalesCount(dateM, "16719", "0")  'הודעות קישורית
            P16719Kishurit_Service = func.GetSalesCount(dateM, "16719", "1")
            Dim Upd As New SqlClient.SqlCommand("SET DATEFORMAT dmy;UPDATE DimDate SET  P16719Kishurit_Sales=@P16719Kishurit_Sales,P16719Kishurit_Service=@P16719Kishurit_Service where DateKey=@dateM", conU)
            Upd.Parameters.Add("@P16719Kishurit_Sales", SqlDbType.Int).Value = P16719Kishurit_Sales
            Upd.Parameters.Add("@P16719Kishurit_Service", SqlDbType.Int).Value = P16719Kishurit_Service
            Upd.Parameters.Add("@dateM", SqlDbType.VarChar).Value = dateM
            conU.Open()
            Upd.CommandType = CommandType.Text
            Upd.ExecuteNonQuery()
            conU.Close()

        End While
        dr_m.Close()
        cmdSelect.Dispose()
        con.Close()



    End Sub

End Class

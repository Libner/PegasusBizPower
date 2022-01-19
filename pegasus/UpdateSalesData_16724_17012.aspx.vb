Public Class UpdateSalesData_16724_17012
    Inherits System.Web.UI.Page

    Public func As New bizpower.cfunc
    Public dr_m As SqlClient.SqlDataReader
    Protected dateM As String
    Protected P16724ContactUs_Sales, P16724ContactUs_Service, P17012ContactUs_Sales, P17012ContactUs_Service As String
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
        'Put user code to initialize the page here
        pYear = Request.QueryString("y")
        pMonth = Request.QueryString("p")
        Dim sql As String

        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        If IsNumeric(pYear) And IsNumeric(pMonth) Then
            ' sql = "SET DATEFORMAT dmy;SELECT  DateKey from DimDate  where year(DateKey)=" & pYear & " and month(DateKey)=" & pMonth
            sql = "SET DATEFORMAT dmy;SELECT  DateKey from DimDate  where year(DateKey)=" & pYear & " and Month(DateKey)=" & pMonth

        Else
            sql = "SET DATEFORMAT dmy;SELECT  DateKey from DimDate  where (DateDiff(dd, DateKey, GetDate()) = 1)"
        End If
        Dim cmdSelect As New SqlClient.SqlCommand(sql, con)

        con.Open()
        dr_m = cmdSelect.ExecuteReader()
        Dim urlF
        While dr_m.Read
            dateM = dr_m("DateKey")
           
            ' Try
            '  P16504 = func.GetSalesCount16504(dateM, "16504") ' טפסי מתעניין
            P16724ContactUs_Sales = func.GetSalesCount(dateM, "16724", "0")

            P16724ContactUs_Service = func.GetSalesCount(dateM, "16724", "1")

            P17012ContactUs_Sales = func.GetSalesCount(dateM, "17012", "0")
            P17012ContactUs_Service = func.GetSalesCount(dateM, "17012", "1")
            '  Response.Write(dateM & "-" & P16724ContactUs_Sales & "-" & P17012ContactUs_Sales & "<BR>")

            Dim Upd As New SqlClient.SqlCommand("SET DATEFORMAT dmy;UPDATE DimDate SET  P16724ContactUs_Sales=@P16724Sales,P16724ContactUs_Service=@P16724Service ,P17012ContactUs_Sales=@P17012Sales,P17012ContactUs_Service=@P17012Service where DateKey=@dateM", conU)
            Upd.Parameters.Add("@P16724Sales", SqlDbType.Int).Value = P16724ContactUs_Sales
            Upd.Parameters.Add("@P16724Service", SqlDbType.Int).Value = P16724ContactUs_Service
            Upd.Parameters.Add("@P17012Sales", SqlDbType.Int).Value = P17012ContactUs_Sales
            Upd.Parameters.Add("@P17012Service", SqlDbType.Int).Value = P17012ContactUs_Service


            Upd.Parameters.Add("@dateM", SqlDbType.VarChar).Value = dateM
            conU.Open()
            Upd.CommandType = CommandType.Text
            Upd.ExecuteNonQuery()
            conU.Close()
            '  Catch ex As Exception
            '   Response.Write(ex.Message.ToString())
            '  Response.Write(Err.Description)
            '  Response.Write(dateM)
            ' Response.Write("<Br>" & sqls & "<BR>")
            '   Response.End()
            ' End Try
        End While
        dr_m.Close()
        cmdSelect.Dispose()
        con.Close()



    End Sub

End Class

Public Class updateSalesData
    Inherits System.Web.UI.Page

    Public func As New bizpower.cfunc
    Public dr_m As SqlClient.SqlDataReader
    Protected dateM As String
    Protected P16719Kishurit_Sales, P16719Kishurit_Service, P16724ContactUs_Sales, P16724ContactUs_Service, P16504 As String




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
        Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        Dim cmdSelect As New SqlClient.SqlCommand("SET DATEFORMAT dmy;SELECT  DateKey from DimDate " & _
        " where (datediff(yy,DateKey,GetDate())<=1) and (datediff(yy,DateKey,GetDate())>=0)", con)

        '" where (datediff(dd,DateKey,GetDate())>=0) and  (datediff(dd,DateKey,GetDate())<=100) ", con)
        'and  (datediff(dd,DateKey,GetDate())<=100)
        cmdSelect.CommandType = CommandType.Text
        con.Open()
        dr_m = cmdSelect.ExecuteReader()
        Dim urlF
        While dr_m.Read
            dateM = dr_m("DateKey")
            Response.Write("d=" & dateM & "<bR>")
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
            ' P16724ContactUs_Sales = 0 '  func.GetSalesCount(dateM, "16724", "0")
            '  P16724ContactUs_Service = 0 'func.GetSalesCount(dateM, "16724", "1")
            'P16504 = func.GetSalesCount16504(dateM, "16504") ' טפסי מתעניין
            sqls = "SET DATEFORMAT dmy;UPDATE DimDate SET P16719Kishurit_Sales=" & P16719Kishurit_Sales & "," & _
            " P16719Kishurit_Service=" & P16719Kishurit_Service & ",P16724ContactUs_Sales=" & P16724ContactUs_Sales & ",P16724ContactUs_Service=" & P16724ContactUs_Service & " where DateKey=" & dateM
            'Dim Upd As New SqlClient.SqlCommand("SET DATEFORMAT dmy;UPDATE DimDate SET P16719Kishurit_Sales=@P16719Kishurit_Sales," & _
            ' " P16504=@P16504, P16719Kishurit_Service=@P16719Kishurit_Service,P16724ContactUs_Sales=@P16724ContactUs_Sales,P16724ContactUs_Service=@P16724ContactUs_Service where DateKey=@dateM", conU)
            Dim Upd As New SqlClient.SqlCommand("SET DATEFORMAT dmy;UPDATE DimDate SET  P16719Kishurit_Sales=@P16719Kishurit_Sales,P16719Kishurit_Service=@P16719Kishurit_Service where DateKey=@dateM", conU)


            Upd.Parameters.Add("@P16719Kishurit_Sales", SqlDbType.Int).Value = P16719Kishurit_Sales
            Upd.Parameters.Add("@P16719Kishurit_Service", SqlDbType.Int).Value = P16719Kishurit_Service
            ' Upd.Parameters.Add("@P16724ContactUs_Sales", SqlDbType.Int).Value = P16724ContactUs_Sales
            'Upd.Parameters.Add("@P16724ContactUs_Service", SqlDbType.Int).Value = P16724ContactUs_Service
            'Upd.Parameters.Add("@P16504", SqlDbType.Int).Value = P16504
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

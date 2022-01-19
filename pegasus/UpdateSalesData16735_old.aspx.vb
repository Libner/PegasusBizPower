Imports System.Data.SqlClient
Public Class UpdateSalesData16735
    Inherits System.Web.UI.Page

    Public func As New bizpower.cfunc
    Public dr_m As SqlClient.SqlDataReader
    Protected dateM As String
    Protected dep As Integer

    Protected P16735, P16735_Bitulim As String
    Public pYear As String





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

        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        Dim strSql As String = "SELECT departmentId FROM dbo.Departments"
        Dim dt As New DataTable
        con.Open()

        Dim dta = New SqlClient.SqlDataAdapter(strSql, con)
        dt.BeginLoadData()
        dta.Fill(dt)
        dt.EndLoadData()
        con.Close()
        '''For i As Integer = 0 To dt.DefaultView.Count - 1
        '''    Response.Write("i=" & dt.DefaultView(i).Item("departmentId"))
        '''Next
        '''Response.End()
        Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        '  Dim cmdSelect As New SqlClient.SqlCommand("SET DATEFORMAT dmy;SELECT  DateKey,ID from DimDate " & _
        '  " where (datediff(yy,DateKey,GetDate())<=4) and (datediff(yy,DateKey,GetDate())>=0)", con)
        'Dim cmdSelect As New SqlClient.SqlCommand("SET DATEFORMAT dmy;SELECT  DateKey,ID from DimDate" & _
        '"  where(DateDiff(dd, DateKey, GetDate()) = 0)", con)
        Dim cmdSelect As New SqlClient.SqlCommand("SET DATEFORMAT dmy;SELECT  DateKey,ID from DimDate where year(DateKey)=" & pYear, con)

        '" where (datediff(dd,DateKey,GetDate())>=0) and  (datediff(dd,DateKey,GetDate())<=100) ", con)
        'and  (datediff(dd,DateKey,GetDate())<=100)
        cmdSelect.CommandType = CommandType.Text
        con.Open()
        dr_m = cmdSelect.ExecuteReader()
        Dim urlF
        While dr_m.Read
            dateM = Day(dr_m("DateKey")) & "/" & Month(dr_m("DateKey")) & "/" & Year(dr_m("DateKey"))
            ID = dr_m("ID")
            '  Response.Write("d=" & dateM & "<bR>")


            For i As Integer = 0 To dt.DefaultView.Count - 1
                dep = dt.DefaultView(i).Item("departmentId")



                ' Response.End()
                'appeal_CallStatus= 1 'פניית שירות
                'else פניית מכירה
                'Dim questionsid, statusId



                '     questionsid = 16719
                '  statusId = 1

                Dim sqls As String

                'Try
                P16735 = func.GetCountSaleByDate(dep, 40660, dateM)
                '  P16735 = func.GetSalesCount16735(dateM, "16735", dep)
                P16735_Bitulim = func.GetSalesCount16735Bitulim(dep, dateM)
                '   Response.Write("1=" & P16735 & ":" & P16735_Bitulim)
                '  Response.End()
                Dim cmdUpdate As New SqlClient.SqlCommand("Departments_Data_update", conU)
                cmdUpdate.CommandType = CommandType.StoredProcedure
                cmdUpdate.Parameters.Add("@P16735", SqlDbType.Int).Value = P16735
                cmdUpdate.Parameters.Add("@P16735_Bitulim", SqlDbType.Int).Value = P16735_Bitulim
                cmdUpdate.Parameters.Add("@departmentId", SqlDbType.Int).Value = dep
                cmdUpdate.Parameters.Add("@DateKey", SqlDbType.DateTime).Value = dateM
                cmdUpdate.Parameters.Add("@Id_DimDate", SqlDbType.Int).Value = ID
                conU.Open()
                cmdUpdate.ExecuteNonQuery()
                cmdUpdate.Dispose()
                conU.Close()

                'Catch ex As Exception
                '    Response.Write(ex.Message.ToString())
                '    Response.Write(Err.Description)
                '    Response.Write(dateM)
                '    Response.Write("<Br>" & sqls & "<BR>")
                '    Response.End()
                'End Try
            Next
        End While
        dr_m.Close()
        cmdSelect.Dispose()
        con.Close()
        Response.Write("end")



    End Sub

End Class

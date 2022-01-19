Imports System.Data.SqlClient
Public Class updateDepartments_Data_Monthly
    Inherits System.Web.UI.Page

        Public func As New bizpower.cfunc
        Public dr_m As SqlClient.SqlDataReader
        Protected dateM As String
        Protected dep As Integer

    'Protected P16735, P16735_Bitulim As String
    Public pYear As String
    Public pMonth, sP16735, sP16735_Bitulim As Integer
    Protected Sales, Sales_Point, KishuritService_Point, Kishuriot, SP16719Kishurit_Service, Scall_Service, ActiveTime_Point, ActiveTime As Decimal







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
        '  Response.Write("GGG")
        '  Response.End()
        'Put user code to initialize the page here
        Session.LCID = 1037
        pYear = Request.QueryString("y")
        If Request.QueryString("p") = "" Then
            pMonth = Month(Now())
        Else
            pMonth = func.fixNumeric(Request.QueryString("p"))
        End If

        If pYear = "" Then
            pYear = Year(Now())
        End If
        ' Response.Write(pYear & "--" & pMonth)
        ' Response.End()
        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        Dim strSql As String = "SELECT departmentId FROM dbo.Departments"
        Dim dt As New DataTable
        con.Open()

        Dim dta = New SqlClient.SqlDataAdapter(strSql, con)
        dt.BeginLoadData()
        dta.Fill(dt)
        dt.EndLoadData()
        con.Close()


        Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        For i As Integer = 0 To dt.DefaultView.Count - 1
            dep = dt.DefaultView(i).Item("departmentId")



            If Not IsNumeric(pMonth) Then
                For pMonth = 1 To 12


                    '   Dim strSales = "select sum(P16735) as sP16735,sum(P16735_Bitulim) as sP16735_Bitulim from Departments_Data  where departmentId=@departmentId and (month(DateKey)=@Dim_Month and year(DateKey)=@DimYear)"
                    '    Response.Write(strSales)

                    Dim cmdSales = New SqlClient.SqlCommand("select sum(ISNULL(P16735,0)) as sP16735,sum(ISNULL(P16735_Bitulim,0)) as sP16735_Bitulim from Departments_Data  where departmentId=@departmentId and (month(DateKey)=@Dim_Month and year(DateKey)=@DimYear)", conU)
                    ' Dim cmdSales As New SqlClient.SqlCommand(strSales, conU)
                    cmdSales.CommandType = CommandType.Text
                    cmdSales.Parameters.Add("@departmentId", SqlDbType.Int).Value = dep
                    cmdSales.Parameters.Add("@Dim_Month", SqlDbType.Char, 2).Value = pMonth
                    cmdSales.Parameters.Add("@DimYear", SqlDbType.Char, 4).Value = pYear
                    conU.Open()
                    Dim dr_m As SqlClient.SqlDataReader = cmdSales.ExecuteReader(CommandBehavior.SingleRow)
                    If dr_m.Read() Then


                        If Not dr_m("sP16735") Is DBNull.Value Then
                            sP16735 = dr_m("sP16735")
                        End If
                        If Not dr_m("sP16735_Bitulim") Is DBNull.Value Then
                            sP16735_Bitulim = dr_m("sP16735_Bitulim")
                        End If
                        ' Response.Write("sP16735=" & sP16735)
                        ' Response.End()


                    End If
                    Sales = sP16735 - sP16735_Bitulim
                    'Response.Write("sP16735=" & sP16735)
                    conU.Close()
                    Dim cmdKishuriot = New SqlClient.SqlCommand("select  sum(P16719Kishurit_Service) as SP16719Kishurit_Service,sum(call_Service) as Scall_Service from  DimDate " & _
                    " where  (month(DateKey)=@Dim_Month and year(DateKey)=@DimYear)", conU)
                    cmdKishuriot.Parameters.Add("@Dim_Month", SqlDbType.Char, 2).Value = pMonth
                    cmdKishuriot.Parameters.Add("@DimYear", SqlDbType.Char, 4).Value = pYear
                    conU.Open()
                    Dim dr_k As SqlClient.SqlDataReader = cmdKishuriot.ExecuteReader(CommandBehavior.SingleRow)
                    If dr_k.Read() Then
                        If Not dr_k("SP16719Kishurit_Service") Is DBNull.Value Then
                            SP16719Kishurit_Service = dr_k("SP16719Kishurit_Service")
                        End If
                        If Not dr_k("Scall_Service") Is DBNull.Value Then
                            Scall_Service = dr_k("Scall_Service")
                        End If
                    End If
                    If Scall_Service > 0 Then
                        Kishuriot = Math.Round((SP16719Kishurit_Service / Scall_Service) * 100, 2)
                    Else
                        Kishuriot = 0
                    End If
                    conU.Close()


                    Dim cmdUpdate As New SqlClient.SqlCommand("Departments_DataMonthly_update", conU)
                    cmdUpdate.CommandType = CommandType.StoredProcedure
                    cmdUpdate.Parameters.Add("@Sales", SqlDbType.Float).Value = Sales  'נטו רישום 
                    cmdUpdate.Parameters.Add("@Sales_Point", SqlDbType.Float).Value = Sales_Point  ' not update
                    cmdUpdate.Parameters.Add("@KishuritService_Point", SqlDbType.Float).Value = KishuritService_Point  ' not update
                    cmdUpdate.Parameters.Add("@Kishuriot", SqlDbType.Float).Value = Kishuriot
                    cmdUpdate.Parameters.Add("@ActiveTime_Point", SqlDbType.Float).Value = ActiveTime_Point  ' not update
                    cmdUpdate.Parameters.Add("@ActiveTime", SqlDbType.Float).Value = ActiveTime
                    cmdUpdate.Parameters.Add("@departmentId", SqlDbType.Int).Value = dep
                    cmdUpdate.Parameters.Add("@Dim_Month", SqlDbType.Char, 2).Value = pMonth
                    cmdUpdate.Parameters.Add("@DimYear", SqlDbType.Char, 4).Value = pYear
                    conU.Open()
                    cmdUpdate.ExecuteNonQuery()
                    cmdUpdate.Dispose()
                    conU.Close()


                Next

            Else
                '''data only one month
                Dim cmdSales = New SqlClient.SqlCommand("select sum(ISNULL(P16735,0)) as sP16735,sum(ISNULL(P16735_Bitulim,0)) as sP16735_Bitulim from Departments_Data  where departmentId=@departmentId and (month(DateKey)=@Dim_Month and year(DateKey)=@DimYear)", conU)
                ' Dim cmdSales As New SqlClient.SqlCommand(strSales, conU)
                cmdSales.CommandType = CommandType.Text
                cmdSales.Parameters.Add("@departmentId", SqlDbType.Int).Value = dep
                cmdSales.Parameters.Add("@Dim_Month", SqlDbType.Char, 2).Value = pMonth
                cmdSales.Parameters.Add("@DimYear", SqlDbType.Char, 4).Value = pYear
                conU.Open()
                Dim dr_m As SqlClient.SqlDataReader = cmdSales.ExecuteReader(CommandBehavior.SingleRow)
                If dr_m.Read() Then


                    If Not dr_m("sP16735") Is DBNull.Value Then
                        sP16735 = dr_m("sP16735")
                    End If
                    If Not dr_m("sP16735_Bitulim") Is DBNull.Value Then
                        sP16735_Bitulim = dr_m("sP16735_Bitulim")
                    End If
                    ' Response.Write("sP16735=" & sP16735)
                    ' Response.End()


                End If
                Sales = sP16735 - sP16735_Bitulim
                'Response.Write("sP16735=" & sP16735)
                conU.Close()
                Dim cmdKishuriot = New SqlClient.SqlCommand("select  sum(P16719Kishurit_Service) as SP16719Kishurit_Service,sum(call_Service) as Scall_Service from  DimDate " & _
                " where  (month(DateKey)=@Dim_Month and year(DateKey)=@DimYear)", conU)
                cmdKishuriot.Parameters.Add("@Dim_Month", SqlDbType.Char, 2).Value = pMonth
                cmdKishuriot.Parameters.Add("@DimYear", SqlDbType.Char, 4).Value = pYear
                conU.Open()
                Dim dr_k As SqlClient.SqlDataReader = cmdKishuriot.ExecuteReader(CommandBehavior.SingleRow)
                If dr_k.Read() Then
                    If Not dr_k("SP16719Kishurit_Service") Is DBNull.Value Then
                        SP16719Kishurit_Service = dr_k("SP16719Kishurit_Service")
                    End If
                    If Not dr_k("Scall_Service") Is DBNull.Value Then
                        Scall_Service = dr_k("Scall_Service")
                    End If
                End If
                If Scall_Service > 0 Then
                    Kishuriot = Math.Round((SP16719Kishurit_Service / Scall_Service) * 100, 2)
                Else
                    Kishuriot = 0
                End If
                conU.Close()


                Dim cmdUpdate As New SqlClient.SqlCommand("Departments_DataMonthly_update", conU)
                cmdUpdate.CommandType = CommandType.StoredProcedure
                cmdUpdate.Parameters.Add("@Sales", SqlDbType.Float).Value = Sales  'נטו רישום 
                cmdUpdate.Parameters.Add("@Sales_Point", SqlDbType.Float).Value = Sales_Point  ' not update
                cmdUpdate.Parameters.Add("@KishuritService_Point", SqlDbType.Float).Value = KishuritService_Point  ' not update
                cmdUpdate.Parameters.Add("@Kishuriot", SqlDbType.Float).Value = Kishuriot
                cmdUpdate.Parameters.Add("@ActiveTime_Point", SqlDbType.Float).Value = ActiveTime_Point  ' not update
                cmdUpdate.Parameters.Add("@ActiveTime", SqlDbType.Float).Value = ActiveTime
                cmdUpdate.Parameters.Add("@departmentId", SqlDbType.Int).Value = dep
                cmdUpdate.Parameters.Add("@Dim_Month", SqlDbType.Char, 2).Value = pMonth
                cmdUpdate.Parameters.Add("@DimYear", SqlDbType.Char, 4).Value = pYear
                conU.Open()
                cmdUpdate.ExecuteNonQuery()
                cmdUpdate.Dispose()
                conU.Close()




            End If
            'Try

            '   Response.Write("1=" & P16735 & ":" & P16735_Bitulim)
            '  Response.End()

            'Catch ex As Exception
            '    Response.Write(ex.Message.ToString())
            '    Response.Write(Err.Description)
            '    Response.Write(dateM)
            '    Response.Write("<Br>" & sqls & "<BR>")
            '    Response.End()
            'End Try
        Next

        con.Close()
        'Response.Write("end")



    End Sub

End Class

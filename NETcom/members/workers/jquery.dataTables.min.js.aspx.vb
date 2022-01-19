Public Class workersTmp
    Inherits System.Web.UI.Page
    Protected arch As Integer
    Protected WithEvents rptWorkers As WebControls.Repeater
    Dim con As New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected OrgId As Integer = 264
    Public func As New bizpower.cfunc
    Protected lang_id As Integer = 1
    Protected dir_obj_var As String = "rtl"
    Protected str_alert As String

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
        arch = 0
        If Request.QueryString("arch") <> "" Then
            If IsNumeric(Request.QueryString("arch")) Then
                arch = Request.QueryString("arch")
            End If
        Else
            arch = 0
        End If
        BindData()

    End Sub
    Private Sub BindData()
        Dim sqlstr As String = "users_site_listGrid"
        Dim cmdSelect As New SqlClient.SqlCommand(sqlstr, con)
        cmdSelect.CommandType = CommandType.StoredProcedure
        'cmdSelect.Parameters.Add("@TourId", SqlDbType.Int).Value = drpTours.SelectedValue
        cmdSelect.Parameters.Add("@OrgId", SqlDbType.Int).Value = OrgId
        cmdSelect.Parameters.Add("@arch", SqlDbType.Int).Value = arch


        con.Open()

        rptWorkers.DataSource = cmdSelect.ExecuteReader()
        rptWorkers.DataBind()
        con.Close()
        con.Dispose()

        'dtgWorkers.DataSource = dt
        'dtgWorkers.DataKeyField = "PFile"
        'dtgWorkers.DataBind()
    End Sub
End Class

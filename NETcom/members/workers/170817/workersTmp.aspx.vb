Public Class workersTmp
    Inherits System.Web.UI.Page
    Protected arch As Integer
    Protected WithEvents rptWorkers, rptTitles As WebControls.Repeater
    Protected OrgId As Integer = 264
    Public func As New bizpower.cfunc
    Protected lang_id As Integer = 1
    Protected dir_obj_var As String = "rtl"
    Protected str_alert As String
    Protected lblImpId As Label
    Protected i As Integer


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
        BindTitle()
        'BindDataPer()
        BindData()

    End Sub
    'Private Sub BindDataPer()
    '    Dim sqlTitleP = "SELECT bar_id, bar_title, bar_title_eng, parent_id, bar_url, bar_order, in_visible FROM bars WHERE (in_use = 1)"
    '    Dim cmdSelect As New SqlClient.SqlCommand(sqlTitleP, con)
    '    cmdSelect.CommandType = CommandType.Text
    '    con.Open()
    '    rptData.DataSource = cmdSelect.ExecuteReader()
    '    rptData.DataBind()
    '    con.Close()
    '    con.Dispose()

    'End Sub
    Private Sub BindTitle()
        Dim con2 As New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        Dim sqlTitle = "Select bar_id,bar_title,parent_id,parent_title from dbo.bar_organizations_table('264') " & _
            " WHERE PARENT_ID IS NOT NULL AND IS_VISIBLE = '1' Order by Parent_Order desc, bar_Order desc "
        Dim cmdSelect As New SqlClient.SqlCommand(sqlTitle, con2)
        cmdSelect.CommandType = CommandType.Text
        con2.Open()
        rptTitles.DataSource = cmdSelect.ExecuteReader()
        rptTitles.DataBind()
        con2.Close()
        con2.Dispose()

    End Sub
    Private Sub BindData()
        Dim con As New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

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

    Private Sub rptWorkers_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptWorkers.ItemDataBound
        Dim con1 As New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

        'If (e.Item.ItemType = ListItemType.Header) Then
        '    Dim rptTitles As Repeater

        '    rptTitles = e.Item.FindControl("rptTitles")
        '    'Dim UserId = e.Item.DataItem("User_Id")
        '    Dim sqlTitle = "SELECT bar_id, bar_title, bar_title_eng, parent_id, bar_url, bar_order, in_use FROM bars WHERE (in_use = 1) order  by bar_id desc"
        '    Dim cmdSelect As New SqlClient.SqlCommand(sqlTitle, con1)
        '    cmdSelect.CommandType = CommandType.Text
        '    con1.Open()
        '    rptTitles.DataSource = cmdSelect.ExecuteReader()
        '    rptTitles.DataBind()
        '    con1.Close()
        '    con1.Dispose()
        'End If


        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then

            lblImpId = e.Item.FindControl("lblImpId")
            Dim ImportanceId As String = func.dbNullFix(e.Item.DataItem("ImportanceId"))
            Dim rptData As Repeater
            rptData = e.Item.FindControl("rptData")
            Dim UserId = e.Item.DataItem("User_Id")
            Dim sqlData = "Select T.bar_id,bar_title,parent_id,parent_title," & _
  " Is_Permision=case when BU.User_Id is not null then BU.is_Visible else 0 end " & _
  "	from dbo.bar_organizations_table('264') as T	 left outer join  bar_users BU on " & _
  " BU.bar_id = T.bar_id And User_Id = @UserId WHERE PARENT_ID IS NOT NULL AND T.IS_VISIBLE = '1'  " & _
  " Order by Parent_Order desc , bar_Order desc"
            Dim cmdSelectU As New SqlClient.SqlCommand(sqlData, con1)
            cmdSelectU.CommandType = CommandType.Text
            cmdSelectU.Parameters.Add("@UserId", SqlDbType.Int).Value = UserId
            con1.Open()
            rptData.DataSource = cmdSelectU.ExecuteReader()
            rptData.DataBind()
            con1.Close()
            con1.Dispose()
            Select Case ImportanceId
                    Case "1"
                        lblImpId.Text = "שוטף"
                    Case "2"
                        lblImpId.Text = "חשיבות נמוכה"
                    Case "3"
                        lblImpId.Text = "חשוב"
                    Case "4"
                        lblImpId.Text = "חשוב מאוד"
                    Case Else
                        lblImpId.Text = ""
                End Select

        End If
    End Sub
End Class

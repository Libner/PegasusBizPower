Public Class orders_check
    Inherits System.Web.UI.Page
    Protected userId, orgID, strLocal, user_name, org_name, perSize, srchPFile, srchService As String
    Protected WithEvents dtgGilboa As WebControls.DataGrid
    Protected cmbUser As HtmlSelect
    Dim con As New System.Data.SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public dir_var, align_var, dir_obj_var, lang_id As String
    Public arrTitles(1), arrButtons(1)
    Public title_count, button_count As Integer
    Protected func As New bizpower.cfunc
    Protected topIncludeU As topInclude


#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here     
        strLocal = "http://" & Trim(Request.ServerVariables("SERVER_NAME"))
        If Len(Trim(Request.ServerVariables("SERVER_PORT"))) > 0 Then
            strLocal = strLocal & ":" & Trim(Request.ServerVariables("SERVER_PORT"))
        End If
        strLocal = strLocal & Application("VirDir") & "/"
        userId = Trim(Request.Cookies("bizpegasus")("UserId"))
        orgID = Trim(Request.Cookies("bizpegasus")("orgId"))
        perSize = Trim(Request.Cookies("bizpegasus")("perSize"))
        lang_id = Trim(Request.Cookies("bizpegasus")("LANGID"))
        org_name = Request.Cookies("bizpegasus")("ORGNAME")
        org_name = HttpContext.Current.Server.UrlDecode(org_name)
        user_name = Trim(Request.Cookies("bizpegasus")("UserName"))
        user_name = HttpContext.Current.Server.UrlDecode(user_name)
        If IsNothing(Request.Cookies("bizpegasus")) Then
            '   Response.Write("1=" & Len(Request.Cookies("bizpegasus")))
            Response.Redirect(strLocal)
        End If
        If Not IsNumeric(Request.Cookies("bizpegasus")("UserId")) Then
            Response.Redirect(strLocal)
        End If

        If IsNumeric(lang_id) = False Or lang_id = "" Then
            lang_id = "1"
        End If
        If lang_id = "2" Then
            dir_var = "rtl" : align_var = "left" : dir_obj_var = "ltr"
        Else
            dir_var = "ltr" : align_var = "right" : dir_obj_var = "rtl"
        End If

        srchPFile = Trim(Request.QueryString("srchPFile"))
        srchService = Trim(Request.QueryString("srchService"))

        If Page.IsPostBack = False Then
            dtgGilboa.Attributes("SortExpression") = "PFile"
            dtgGilboa.Attributes("SortASC") = "Yes"
            BindData()
        End If

    End Sub

    Private Sub BindData()

        Dim sortExpr As String = ""
        If Len(dtgGilboa.Attributes("SortExpression")) > 0 Then
            sortExpr = "[" & dtgGilboa.Attributes("SortExpression") & "] " & IIf(dtgGilboa.Attributes("SortASC") = "No", " DESC", "")
        Else
            sortExpr = "PFile"
        End If

        Dim sqlstr As String = "SELECT [PFile], [Service Name], [Pax Num] FROM gilboa WHERE (1=1)"
        If Len(srchPFile) > 0 Then
            sqlstr = sqlstr & " AND ([PFile] LIKE '%" & func.sFix(srchPFile) & "%') "
        End If
        If Len(srchService) > 0 Then
            sqlstr = sqlstr & " AND ([Service Name] LIKE '%" & func.sFix(srchService) & "%') "
        End If
        sqlstr = sqlstr & " ORDER BY " & sortExpr & ""
        Dim cmdSelect As New SqlClient.SqlCommand(sqlstr, con)
        cmdSelect.CommandType = CommandType.Text
        con.Open()
        Dim da As New SqlClient.SqlDataAdapter(cmdSelect)
        Dim dt As New DataTable
        dt.BeginLoadData()
        da.Fill(dt)
        dt.EndLoadData()
        con.Close()

        dtgGilboa.DataSource = dt
        dtgGilboa.DataKeyField = "PFile"
        dtgGilboa.DataBind()

        dtgGilboa.Visible = Not (dtgGilboa.Items.Count = 0)

        cmdSelect = New SqlClient.SqlCommand("SELECT User_ID, (FIRSTNAME + ' ' + LASTNAME) UserName FROM Users WHERE (ORGANIZATION_ID = " & orgID & ") AND (ACTIVE = 1) ORDER BY FIRSTNAME + ' ' + LASTNAME", con)
        cmdSelect.CommandType = CommandType.Text
        con.Open()
        cmbUser.DataSource = cmdSelect.ExecuteReader()
        cmbUser.DataTextField = "UserName"
        cmbUser.DataValueField = "User_ID"
        cmbUser.DataBind()
        cmdSelect.Dispose() : con.Close()

        cmbUser.Items.Insert(0, New ListItem("כל העובדים", 0))

    End Sub

    Sub dtgGilboa_PageChanger(ByVal Source As Object, _
  ByVal E As DataGridPageChangedEventArgs) Handles dtgGilboa.PageIndexChanged
        ' Set the CurrentPageIndex before binding the grid 
        dtgGilboa.CurrentPageIndex = E.NewPageIndex

        BindData()
    End Sub

    Private Sub dtgGilboa_SortCommand(ByVal source As Object, _
ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) _
Handles dtgGilboa.SortCommand
        Dim strSort = dtgGilboa.Attributes("SortExpression")
        Dim strASC = dtgGilboa.Attributes("SortASC")

        dtgGilboa.Attributes("SortExpression") = e.SortExpression
        dtgGilboa.Attributes("SortASC") = "Yes"

        If e.SortExpression = strSort Then
            If strASC = "Yes" Then
                dtgGilboa.Attributes("SortASC") = "No"
            Else
                dtgGilboa.Attributes("SortASC") = "Yes"
            End If
        End If

        BindData()

    End Sub

    Private Sub DisplaySortOrderImages(ByVal sortExpression As String, ByVal dgItem As DataGridItem)
        Dim sortColumns As String() = sortExpression.Split(",".ToCharArray())
        For i As Integer = 0 To dgItem.Cells.Count - 1
            If dgItem.Cells(i).Controls.Count > 0 AndAlso TypeOf dgItem.Cells(i).Controls(0) Is LinkButton Then

                Dim column As String = DirectCast(dgItem.Cells(i).Controls(0), LinkButton).CommandArgument

                If Trim(dtgGilboa.Attributes("SortExpression")) = Trim(column) Then
                    If dtgGilboa.Attributes("SortASC") = "Yes" Then
                        Dim imgUp As New WebControls.Image
                        imgUp.ImageUrl = Application("VirDir") & "/NETcom/images/arrow_up.gif"
                        dgItem.Cells(i).Controls.Add(imgUp)
                    ElseIf dtgGilboa.Attributes("SortASC") = "No" Then
                        Dim imgDown As New WebControls.Image
                        imgDown.ImageUrl = Application("VirDir") & "/NETcom/images/arrow_down.gif"
                        dgItem.Cells(i).Controls.Add(imgDown)
                    End If

                End If
            End If
        Next

    End Sub

    Private Sub dtgGilboa_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dtgGilboa.ItemDataBound
        If e.Item.ItemType = ListItemType.Header Then
            If dtgGilboa.Attributes("SortExpression") <> "" Then
                DisplaySortOrderImages(dtgGilboa.Attributes("SortExpression"), e.Item)
            End If
        End If
    End Sub

End Class
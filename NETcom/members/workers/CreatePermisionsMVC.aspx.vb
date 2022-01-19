Imports System.Data.SqlClient
Public Class CreatePermisionsMVC
    Inherits System.Web.UI.Page
    Protected func As New bizpower.cfunc
    Protected UserId As Integer
    Protected UserName As String
    Dim con, conBar As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conAction As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))
    Dim conAction1 As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))

    Protected WithEvents rptSubSiteScreen, rptScreenAction As Repeater
    Protected WithEvents rptSubSites As DataList
    Public dtDep As New DataTable
    Dim primKeydtDep(0) As Data.DataColumn
    Protected WithEvents Button1, Button2 As System.Web.UI.WebControls.Button
    '  Protected WithEvents seldep As System.Web.UI.HtmlControls.HtmlSelect


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
        Button1.Text = "אישור"
      
        'Response.Write("end")
        'For i = 1 To Request.Form.Count
        '    ' Response.Write(Request.Form(i) & "<BR>")
        '    Response.Write("i=" & i & "<BR>")

        'Next
        '   Response.End()

        If IsNumeric(Request("Id")) Then
            UserId = Request("Id")

            Dim cmdSelect = New SqlCommand("SELECT  FIRSTNAME + char(32) +LASTNAME as UserName from Users where User_Id=@UserId", con)
            cmdSelect.Parameters.Add("@UserId", SqlDbType.Int).Value = UserId
            con.Open()
            UserName = cmdSelect.ExecuteScalar
            con.Close()

            If Not IsPostBack Then
                ' getData()
                SubSites()
            End If
            If Page.IsPostBack Then
                '  saveData()
            End If
        Else
            Response.Write("No userId")

        End If
    End Sub
    Sub SubSites()
        Dim cmdSelect1 As SqlClient.SqlCommand
        Dim strsql = "Select SubSitesId,SubSitesName FROM SubSites  order by SubSitesOrder"
        cmdSelect1 = New SqlClient.SqlCommand(strsql, con)
        cmdSelect1.CommandType = CommandType.Text
        con.Open()
        rptSubSites.DataSource = cmdSelect1.ExecuteReader
        rptSubSites.DataBind()
        con.Close()
    End Sub
    Private Sub rptSubSites_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataListItemEventArgs) Handles rptSubSites.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim SubSitesId As Integer = e.Item.DataItem("SubSitesId")

            '  Dim rptSubSiteScreen As Repeater
            'Dim plhdep As PlaceHolder
            'Dim seldep As HtmlSelect
            'seldep = e.Item.FindControl("seldep")
            '-------------------
            'Dim conDep As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))

            'Dim cmdSelectDep = New SqlCommand("Select departmentId, departmentName,PriorityLevel,dbo.IsVisibleMVCScreenDepartment(departmentId," & UserId & ") as is_visible from Departments order by departmentName", conDep)
            'conDep.Open()
            'Dim ad As New SqlClient.SqlDataAdapter
            'ad.SelectCommand = cmdSelectDep
            'ad.Fill(dtDep)
            'conDep.Close()
            'primKeydtDep(0) = dtDep.Columns("departmentId")
            'dtDep.PrimaryKey = primKeydtDep
            '  seldep.Items.Clear()
            '  Dim list1 As New ListItem("הכל", "0")
            '  seldep.Items.Add(list1)
            'For i As Integer = 0 To dtDep.Rows.Count - 1
            '    Dim list As New ListItem(dtDep.Rows(i)("departmentName"), dtDep.Rows(i)("departmentId"))
            '    seldep.Items.Add(list)
            '    If dtDep.Rows(i)("is_visible") = "1" Then
            '        seldep.Items.Item(i).Selected = True
            '    End If
            'Next
            '----------------------
            'plhdep = e.Item.FindControl("plhdep")
            'If SubSitesId = 2 Then
            '    plhdep.Visible = True
            'End If
            Dim rptSubSiteScreen As Repeater = DirectCast(e.Item.FindControl("rptSubSites"), Repeater)

            rptSubSiteScreen = e.Item.FindControl("rptSubSiteScreen")
            '   Response.Write(SubSitesId & "<BR>")
            Dim cmdSelect = New SqlCommand("Select SubSitesId,ScreenId,ScreenName, dbo.IsVisibleMVCScreenPermission(ScreenId," & UserId & ") as is_visible FROM SubSitesScreens where SubSitesId=@SubSitesId order by ScreenOrder", conAction)
            cmdSelect.Parameters.Add("@SubSitesId", SqlDbType.Int).Value = CInt(SubSitesId)
            conAction.Open()
            Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

            rptSubSiteScreen.DataSource = dr
            rptSubSiteScreen.DataBind()
            conAction.Close()
        End If
    End Sub
    Public Sub CreaterptScreenAction(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs)
        ' Response.Write("rptScreenAction")

        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim ScreenId As Integer = e.Item.DataItem("ScreenId")

            Dim rptScreenAction As Repeater
            rptScreenAction = e.Item.FindControl("rptScreenAction")
            Dim cmdSelect = New SqlCommand("SELECT SubSitesId,ActionName,*,dbo.IsVisibleMVCScreenActionPermission(ActionId," & UserId & ") as is_visible  FROM  SubSitesScreenAction left join SubSitesScreens on SubSitesScreens.ScreenId=SubSitesScreenAction.ScreenId WHERE(SubSitesScreenAction.ScreenId = @ScreenId) ORDER BY ActionOrder", conAction1)
            'DISTINCT Permissions.ScreenId, Permissions.ActionId, Permissions.Parent_Id, Permissions.Permission_Name,dbo.IsVisiblePermissions(Permissions.ActionId," & UserId & ") as is_visible, Permissions.Bar_Order
            cmdSelect.Parameters.Add("@ScreenId", SqlDbType.Int).Value = CInt(ScreenId)
            conAction1.Open()
            Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

            If dr.HasRows Then
                '  Response.Write(PageSize.SelectedValue)
                '  Response.End()
                rptScreenAction.DataSource = dr
                rptScreenAction.DataBind()

            End If

            conAction1.Close()

        End If
    End Sub
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        saveData()
        Dim cScript As String

        cScript = "<script language='javascript'>self.close(); </script>"
        RegisterStartupScript("ReloadScrpt", cScript)

    End Sub
    Public Sub SaveData()
        Dim sqlstr As String
        Dim cmd As New SqlClient.SqlCommand
        Response.Write("UserId=" & UserId & "<br>")
        Dim cmdDelete As New SqlClient.SqlCommand("Delete FROM PermissionsMVC WHERE  USER_ID = @UserId;", con)
        cmdDelete.CommandType = CommandType.Text
        cmdDelete.Parameters.Add("@UserId", SqlDbType.Int).Value = UserId
        con.Open()
        '  Try
        cmdDelete.ExecuteNonQuery()
        con.Close()

        Dim i As Integer
        Response.Write("saveData")
        '     Response.Write("is_visibleL=" & Request.Form("is_visibleL") & "<BR>")
        '     Response.Write("hghhhhh" & "<BR>")

        '   Response.Write("is_visible =" & Request.Form("is_visible") & "<BR>")
        Dim pArr, pSubArr As Array
        pArr = Request("is_visible").Split(",")
        For i = 0 To UBound(pArr)
            '  Response.Write(i & "=" & pArr(i) & "<Br>")
            pSubArr = pArr(i).Split("!")
            '  Response.Write(pSubArr(0) & "-" & pSubArr(1) & "-" & pSubArr(2) & "<BR>")
            con.Open()
            sqlstr = "Insert Into PermissionsMVC (User_Id,ScreenId,SubSitesId,ActionId,Permission) values (" & UserId & "," & pSubArr(1) & "," & pSubArr(2) & "," & pSubArr(0) & ",'1')"
            cmd = New System.Data.SqlClient.SqlCommand(sqlstr, con)
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()
            con.Close()
        Next

    End Sub

End Class

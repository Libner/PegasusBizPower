Imports System.Data.SqlClient
Public Class AddWorkerMVC
    Inherits System.Web.UI.Page
    Protected func As New bizpower.cfunc
    Protected UserId As Integer
    Protected UserName As String
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))
    Dim conAction As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))
    Dim conAction1 As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))

    Protected WithEvents rptSubSiteScreen, rptScreenAction As Repeater
    Protected WithEvents rptSubSites As DataList
    Public dtDep As New DataTable
    Dim primKeydtDep(0) As Data.DataColumn
    Protected WithEvents Button1, Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents seldep As System.Web.UI.HtmlControls.HtmlSelect


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
    

        If IsNumeric(Request("Id")) Then
            UserId = Request("Id")
            Dim cmdSelect = New SqlCommand("SELECT  FIRSTNAME + char(32) +LASTNAME as UserName from Users where User_Id=@UserId", con)
            cmdSelect.Parameters.Add("@UserId", SqlDbType.Int).Value = UserId
            con.Open()
            UserName = cmdSelect.ExecuteScalar
            con.Close()
            If Not IsPostBack Then
                SubSites()

            End If

            If Page.IsPostBack Then
                saveData()
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
    'Sub SubSitesScreens(ByVal SubSitesId)
    '    Dim cmdSelect1 As SqlClient.SqlCommand
    '    Dim strsql = "Select ScreenId,ScreenName FROM SubSitesScreens where SubSitesId=@SubSitesId order by ScreenOrder"
    '    cmdSelect1 = New SqlClient.SqlCommand(strsql, con)
    '    cmdSelect1.Parameters.Add("@SubSitesId", SqlDbType.Int).Value = SubSitesId
    '    cmdSelect1.CommandType = CommandType.Text
    '    con.Open()
    '    rptSubSite1.DataSource = cmdSelect1.ExecuteReader
    '    rptSubSite1.DataBind()
    '    con.Close()
    'End Sub



    Private Sub rptSubSites_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataListItemEventArgs) Handles rptSubSites.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim SubSitesId As Integer = e.Item.DataItem("SubSitesId")
            '  Dim rptSubSiteScreen As Repeater
            Dim plhdep As PlaceHolder
            Dim seldep As HtmlSelect
            seldep = e.Item.FindControl("seldep")
            '-------------------
            Dim conDep As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))

            Dim cmdSelectDep = New SqlCommand("Select departmentId, departmentName,PriorityLevel,dbo.IsVisibleMVCScreenDepartment(departmentId," & UserId & ") as is_visible from Departments order by departmentName", conDep)
            conDep.Open()
            Dim ad As New SqlClient.SqlDataAdapter
            ad.SelectCommand = cmdSelectDep
            ad.Fill(dtDep)
            conDep.Close()
            primKeydtDep(0) = dtDep.Columns("departmentId")
            dtDep.PrimaryKey = primKeydtDep
            seldep.Items.Clear()
            '  Dim list1 As New ListItem("הכל", "0")
            '  seldep.Items.Add(list1)
            For i As Integer = 0 To dtDep.Rows.Count - 1
                Dim list As New ListItem(dtDep.Rows(i)("departmentName"), dtDep.Rows(i)("departmentId"))
                seldep.Items.Add(list)
                If dtDep.Rows(i)("is_visible") = "1" Then
                    seldep.Items.Item(i).Selected = True
                End If
            Next
            '----------------------
            plhdep = e.Item.FindControl("plhdep")
            If SubSitesId = 2 Then
                plhdep.Visible = True
            End If
            Dim rptSubSiteScreen As Repeater = DirectCast(e.Item.FindControl("rptSubSites"), Repeater)

            rptSubSiteScreen = e.Item.FindControl("rptSubSiteScreen")
            '   Response.Write(SubSitesId & "<BR>")
            Dim cmdSelect = New SqlCommand("Select ScreenId,ScreenName, dbo.IsVisibleMVCScreen(ScreenId," & UserId & ") as is_visible FROM SubSitesScreens where SubSitesId=@SubSitesId order by ScreenOrder", conAction)
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
            Dim cmdSelect = New SqlCommand("SELECT ActionName,*,dbo.IsVisibleMVCScreenAction(ActionId," & UserId & ") as is_visible  FROM  SubSitesScreenAction WHERE(ScreenId = @ScreenId) ORDER BY ActionOrder", conAction1)
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
    Public Sub saveData()
        Dim conUpd As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim cmd As New SqlClient.SqlCommand
        Dim i As Integer = 0
        Dim barVisible As String
        Dim ActionId, ScreenId As Integer
        ' Dim cmdSelect = New SqlCommand

        Dim cmdDelete As New SqlClient.SqlCommand("Delete FROM SubSitesUsers_ScreenAction WHERE  USERID = @UserId", con)
        cmdDelete.CommandType = CommandType.Text
        cmdDelete.Parameters.Add("@UserId", SqlDbType.Int).Value = UserId
        con.Open()
        cmdDelete.ExecuteNonQuery()
        con.Close()
        '  Response.End()

        Dim dr As SqlDataReader

        Dim cmdSelect As New SqlCommand("Select ScreenId,ActionId from SubSitesScreenAction  Order by ScreenId, ActionOrder", con)
        con.Open()
        dr = cmdSelect.ExecuteReader()

        While dr.Read()

            ScreenId = Trim(dr("ScreenId"))
            ActionId = Trim(dr("ActionId"))

            '     Response.Write("b=" & Request.Form("is_visible" & ActionId) & "<BR>")
            If Trim(Request.Form("is_visible" & ActionId)) <> "" Then
                ' Response.Write(ActionId & ":" & Request.Form("is_visible" & ActionId) & "<BR>")
                barVisible = Request.Form("is_visible" & ActionId)
            End If

            '''''        '''    Response.Write("fff")
            If Trim(barVisible) = "on" Then
                barVisible = 1
                '  Response.Write("barVisible=" & barVisible & "<BR>")
            Else
                barVisible = 0
            End If
            conUpd.Open()
            Dim sqlstr = "Insert Into SubSitesUsers_ScreenAction (UserId,ScreenId,ActionId,ActionStatus) values (" & UserId & "," & ScreenId & "," & ActionId & ",'" & barVisible & "')"
            cmd = New System.Data.SqlClient.SqlCommand(sqlstr, conUpd)
            cmd.CommandType = CommandType.Text
            cmd.ExecuteNonQuery()
            conUpd.Close()
            '   Response.Write(sqlstr & "<BR>")
            '''''        '  Response.End()
        End While
        con.Close()
        dr = Nothing
        'Response.End()
        Response.Write("seldep=" & Request("rptSubSites$ctl01$seldep"))
        conUpd.Open()
        Dim sqlstrDel = "Delete from SubSitesUsers_ScreenDepartments where  UserId=" & UserId
        cmd = New System.Data.SqlClient.SqlCommand(sqlstrDel, conUpd)
        cmd.CommandType = CommandType.Text
        cmd.ExecuteNonQuery()
        conUpd.Close()



        If Len(Request("rptSubSites$ctl01$seldep")) > 0 Then
            '   Dim seldep = Request("rptSubSites$ctl01$seldep")
            Dim depArr As Array
            depArr = Request("rptSubSites$ctl01$seldep").Split(",")
     
            For i = 0 To UBound(depArr)
                conUpd.Open()
                Dim sqlstrIns = "Insert Into SubSitesUsers_ScreenDepartments (UserId,ScreenId,DepartmentId,Departments_Status) values (" & UserId & ",2," & depArr(i) & ",1)"
                cmd = New System.Data.SqlClient.SqlCommand(sqlstrIns, conUpd)
                cmd.CommandType = CommandType.Text
                cmd.ExecuteNonQuery()
                conUpd.Close()

            Next

        End If
    End Sub
    'Sub GetDep()

    '    'If Len(arrPermissionsByDepartmentId) > 0 Then

    '    'End If
    '    Dim conDep As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))

    '    Dim cmdSelect = New SqlCommand("Select departmentId, departmentName,PriorityLevel,dbo.IsVisibleMVCScreenDepartment(departmentId," & UserId & ") as is_visible from Departments order by departmentName", conDep)
    '    conDep.Open()
    '    Dim ad As New SqlClient.SqlDataAdapter
    '    ad.SelectCommand = cmdSelect
    '    ad.Fill(dtDep)
    '    conDep.Close()
    '    primKeydtDep(0) = dtDep.Columns("departmentId")
    '    dtDep.PrimaryKey = primKeydtDep
    '    seldep.Items.Clear()
    '    '  Dim list1 As New ListItem("הכל", "0")
    '    '  seldep.Items.Add(list1)
    '    For i As Integer = 0 To dtDep.Rows.Count - 1
    '        Dim list As New ListItem(dtDep.Rows(i)("is_visible"), dtDep.Rows(i)("departmentId"))
    '        'If (Request("seldep") > 0 And Request("seldep") = dtDep.Rows(i)("departmentId")) Then
    '        '   Response.Write("dep=" & dep & "<BR>")
    '        '  Response.Write("dep=" & dtDep.Rows(i)("departmentId"))

    '        seldep.Items.Add(list)
    '        If dtDep.Rows(i)("IsVisibleMVCScreenDepartment") = 1 Then

    '            seldep.Items.Item(dtDep.Rows(i)("departmentId")).Selected = True
    '        End If
    '    Next
    '    '  Response.Write("arrPermissionsByDepartmentId=" & arrPermissionsByDepartmentId & "--")
    '    '''   Response.End()
    '    'If Len(arrPermissionsByDepartmentId) > 0 And arrPermissionsByDepartmentId <> "0" Then
    '    '  Response.Write(">0")
    '    '  Response.End()

    '    'Dim parts As String() = arrPermissionsByDepartmentId.Split(",")
    '    'For Each part As String In parts
    '    '    seldep.Items.FindByValue(part).Selected = True
    '    'Next
    '    'Else
    '    '    Response.Write("0")
    '    '   End If

    '    'Response.Write("dep=" & dep)
    '    ' Response.End()


    'End Sub



End Class

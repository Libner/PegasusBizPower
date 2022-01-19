Imports System.Data.SqlClient
Public Class CreatePermisions
    Inherits System.Web.UI.Page
    Dim con, conBar As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected func As New bizpower.cfunc
    Protected UserId As Integer
    Protected UserName As String
    Public dtDep As New DataTable
    Dim primKeydtDep(0) As Data.DataColumn
    Protected WithEvents seldep As System.Web.UI.HtmlControls.HtmlSelect
    Protected WithEvents Button1, Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents rptParent, rpt90001, rpt90002, rpt90003, rpt90004, rpt90005, rpt90006, rpt90007, rpt90097 As Repeater
    Protected arrPermissionsByDepartmentId As String


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
            arrPermissionsByDepartmentId = func.PermissionsByDepartmentId(UserId)
            Dim cmdSelect = New SqlCommand("SELECT  FIRSTNAME + char(32) +LASTNAME as UserName from Users where User_Id=@UserId", con)
            cmdSelect.Parameters.Add("@UserId", SqlDbType.Int).Value = UserId
            con.Open()
            UserName = cmdSelect.ExecuteScalar
            con.Close()
        
            If Not IsPostBack Then
                Get90001()
                Get90002()
                Get90003()
                Get90004()
                Get90005()
                Get90006()
                Get90007()
                Get90097()
                GetDep()
              
                GetData()
              
            End If
            If Page.IsPostBack Then
                saveData()
            End If
        Else
            Response.Write("No userId")

        End If
    End Sub
    Sub Get90007()
        Dim cmdSelectParent = New SqlCommand
        ' Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=90007  order by bar_order", con)
        If Len(arrPermissionsByDepartmentId) > 0 Then
            'cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
            '             " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
            '             " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90007 AND  UserPermissions.User_Id=" & UserId & " ) ORDER BY Permissions.Bar_Order", con)
            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
                        " dbo.IsVisiblePermissions(Permissions.Bar_Id," & UserId & ") as is_visible, Permissions.Bar_Order FROM  Permissions " & _
                        " WHERE(Permissions.Parent_Id = 90007) ORDER BY Permissions.Bar_Order", con)
        Else
            '           cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
            '                   " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
            '                " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90007) ORDER BY Permissions.Bar_Order", con)

            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name,0 AS is_visible, " & _
                           "  Permissions.Bar_Order FROM  Permissions  WHERE(Permissions.Parent_Id = 90007) ORDER BY Permissions.Bar_Order", con)

        End If

        con.Open()
        Dim dr As SqlDataReader = cmdSelectParent.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rpt90007.DataSource = dr
            rpt90007.DataBind()
        End If
        con.Close()

    End Sub
    Sub Get90006()
        '  Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=90006  order by bar_order", con)
        'Dim cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
        '" UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
        '" Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90006  AND  UserPermissions.User_Id=" & UserId & " ) ORDER BY Permissions.Bar_Order", con)
        Dim cmdSelectParent = New SqlCommand
        ' Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=90007  order by bar_order", con)
        If Len(arrPermissionsByDepartmentId) > 0 Then
            'cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
            '             " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
            '             " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90006 AND  UserPermissions.User_Id=" & UserId & " ) ORDER BY Permissions.Bar_Order", con)
            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
                        " dbo.IsVisiblePermissions(Permissions.Bar_Id," & UserId & ") as is_visible, Permissions.Bar_Order FROM  Permissions " & _
                        " WHERE(Permissions.Parent_Id = 90006) ORDER BY Permissions.Bar_Order", con)
        Else
            '           cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
            '                   " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
            '                " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90007) ORDER BY Permissions.Bar_Order", con)

            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name,0 AS is_visible, " & _
                           "  Permissions.Bar_Order FROM  Permissions  WHERE(Permissions.Parent_Id = 90006) ORDER BY Permissions.Bar_Order", con)

        End If
        con.Open()
        Dim dr As SqlDataReader = cmdSelectParent.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rpt90006.DataSource = dr
            rpt90006.DataBind()
        End If
        con.Close()

    End Sub
    Sub Get90005()
        'Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=90005  order by bar_order", con)
        'Dim cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
        '     " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
        '     " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90005  AND  UserPermissions.User_Id=" & UserId & " ) ORDER BY Permissions.Bar_Order", con)
        Dim cmdSelectParent = New SqlCommand
        ' Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=90007  order by bar_order", con)
        If Len(arrPermissionsByDepartmentId) > 0 Then
            'cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
            '             " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
            '             " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90005 AND  UserPermissions.User_Id=" & UserId & " ) ORDER BY Permissions.Bar_Order", con)
            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
                        " dbo.IsVisiblePermissions(Permissions.Bar_Id," & UserId & ") as is_visible, Permissions.Bar_Order FROM  Permissions " & _
                        " WHERE(Permissions.Parent_Id = 90005) ORDER BY Permissions.Bar_Order", con)
        Else
            '           cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
            '                   " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
            '                " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90007) ORDER BY Permissions.Bar_Order", con)

            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name,0 AS is_visible, " & _
                           "  Permissions.Bar_Order FROM  Permissions  WHERE(Permissions.Parent_Id = 90005) ORDER BY Permissions.Bar_Order", con)

        End If
        con.Open()
        Dim dr As SqlDataReader = cmdSelectParent.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rpt90005.DataSource = dr
            rpt90005.DataBind()
        End If
        con.Close()

    End Sub
    Sub Get90004()
        '' Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=90004  order by bar_order", con)
        'Dim cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
        '     " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions LEFT JOIN  UserPermissions ON " & _
        '     " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90004  AND  UserPermissions.User_Id=" & UserId & " ) ORDER BY Permissions.Bar_Order", con)
        Dim cmdSelectParent = New SqlCommand
        ' Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=90007  order by bar_order", con)
        If Len(arrPermissionsByDepartmentId) > 0 Then
            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
                             " dbo.IsVisiblePermissions(Permissions.Bar_Id," & UserId & ") as is_visible, Permissions.Bar_Order FROM  Permissions " & _
                             " WHERE(Permissions.Parent_Id = 90004) ORDER BY Permissions.Bar_Order", con)
        Else
            '           cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
            '                   " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
            '                " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90007) ORDER BY Permissions.Bar_Order", con)

            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name,0 AS is_visible, " & _
                           "  Permissions.Bar_Order FROM  Permissions  WHERE(Permissions.Parent_Id = 90004) ORDER BY Permissions.Bar_Order", con)

        End If
        con.Open()
        Dim dr As SqlDataReader = cmdSelectParent.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rpt90004.DataSource = dr
            rpt90004.DataBind()
        End If
        con.Close()

    End Sub
    Sub Get90003()
        ' Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=90003  order by bar_order", con)
        'Dim cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
        '     " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions LEFT JOIN  UserPermissions ON " & _
        '     " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90003  AND  UserPermissions.User_Id=" & UserId & " )ORDER BY Permissions.Bar_Order", con)
        Dim cmdSelectParent = New SqlCommand
        ' Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=90007  order by bar_order", con)
        If Len(arrPermissionsByDepartmentId) > 0 Then
            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
                             " dbo.IsVisiblePermissions(Permissions.Bar_Id," & UserId & ") as is_visible, Permissions.Bar_Order FROM  Permissions " & _
                             " WHERE(Permissions.Parent_Id = 90003) ORDER BY Permissions.Bar_Order", con)
        Else
            '           cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
            '                   " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
            '                " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90007) ORDER BY Permissions.Bar_Order", con)

            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name,0 AS is_visible, " & _
                           "  Permissions.Bar_Order FROM  Permissions  WHERE(Permissions.Parent_Id = 90003) ORDER BY Permissions.Bar_Order", con)

        End If

        con.Open()
        Dim dr As SqlDataReader = cmdSelectParent.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rpt90003.DataSource = dr
            rpt90003.DataBind()
        End If
        con.Close()

    End Sub
    Sub Get90002()
        ''  Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=90002  order by bar_order", con)
        'Dim cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
        '     " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
        '     " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90002  AND  UserPermissions.User_Id=" & UserId & " )ORDER BY Permissions.Bar_Order", con)
        Dim cmdSelectParent = New SqlCommand
        ' Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=90007  order by bar_order", con)
        If Len(arrPermissionsByDepartmentId) > 0 Then
            'cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
            '             "  dbo.IsVisiblePermissions(Permissions.Bar_Id," & UserId & ") as is_visible, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
            '             " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90002 ) ORDER BY Permissions.Bar_Order", con)

            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
                            " dbo.IsVisiblePermissions(Permissions.Bar_Id," & UserId & ") as is_visible, Permissions.Bar_Order FROM  Permissions " & _
                            " WHERE(Permissions.Parent_Id = 90002) ORDER BY Permissions.Bar_Order", con)
        Else
            '           cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
            '                   " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
            '                " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90007) ORDER BY Permissions.Bar_Order", con)

            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name,0 AS is_visible, " & _
                           "  Permissions.Bar_Order FROM  Permissions  WHERE(Permissions.Parent_Id = 90002) ORDER BY Permissions.Bar_Order", con)

        End If
        con.Open()
        'Response.Write(cmdSelectParent.commandtext)
        Dim dr As SqlDataReader = cmdSelectParent.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rpt90002.DataSource = dr
            rpt90002.DataBind()
        End If
        con.Close()

    End Sub
    Sub Get90001()
        ''Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=90001 order by bar_order", con)
        'Dim cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
        '     " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
        '     " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90001  AND  UserPermissions.User_Id=" & UserId & " )ORDER BY Permissions.Bar_Order", con)
        Dim cmdSelectParent = New SqlCommand
        ' Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=90007  order by bar_order", con)
        If Len(arrPermissionsByDepartmentId) > 0 Then
            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
                         " dbo.IsVisiblePermissions(Permissions.Bar_Id," & UserId & ") as is_visible, Permissions.Bar_Order FROM  Permissions " & _
                         " WHERE(Permissions.Parent_Id = 90001) ORDER BY Permissions.Bar_Order", con)
        Else
            '           cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
            '                   " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
            '                " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90007) ORDER BY Permissions.Bar_Order", con)

            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name,0 AS is_visible, " & _
                           "  Permissions.Bar_Order FROM  Permissions  WHERE(Permissions.Parent_Id = 90001) ORDER BY Permissions.Bar_Order", con)

        End If
        con.Open()
        Dim dr As SqlDataReader = cmdSelectParent.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rpt90001.DataSource = dr
            rpt90001.DataBind()
        End If
        con.Close()

    End Sub
    Sub GetData()
         Dim cmdSelectParent = New SqlCommand
        'Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=0", con)
        'Dim cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
        '" UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
        '" Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 0) ORDER BY Permissions.Bar_Order", con)
        If Len(arrPermissionsByDepartmentId) > 0 And arrPermissionsByDepartmentId <> "0" Then
            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
         " dbo.IsVisiblePermissions(Permissions.Bar_Id," & UserId & ") as is_visible,Permissions.Bar_Order FROM  Permissions  WHERE (Permissions.Parent_Id = 0) ORDER BY Permissions.Bar_Order", con)

        Else
            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
                     " 0 as is_visible,Permissions.Bar_Order FROM  Permissions  WHERE (Permissions.Parent_Id = 0) ORDER BY Permissions.Bar_Order", con)

        End If
        con.Open()
        Dim dr As SqlDataReader = cmdSelectParent.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rptParent.DataSource = dr
            rptParent.DataBind()
        End If
        con.Close()

    End Sub
    Sub Get90097()
        Dim cmdSelectParent = New SqlCommand
        ' Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=90007  order by bar_order", con)
        If Len(arrPermissionsByDepartmentId) > 0 Then
            'cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
            '             " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
            '             " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90007 AND  UserPermissions.User_Id=" & UserId & " ) ORDER BY Permissions.Bar_Order", con)
            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
                        " dbo.IsVisiblePermissions(Permissions.Bar_Id," & UserId & ") as is_visible, Permissions.Bar_Order FROM  Permissions " & _
                        " WHERE(Permissions.Parent_Id = 90097) ORDER BY Permissions.Bar_Order", con)
        Else
            '           cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
            '                   " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
            '                " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = 90007) ORDER BY Permissions.Bar_Order", con)

            cmdSelectParent = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name,0 AS is_visible, " & _
                           "  Permissions.Bar_Order FROM  Permissions  WHERE(Permissions.Parent_Id = 90097) ORDER BY Permissions.Bar_Order", con)

        End If

        con.Open()
        Dim dr As SqlDataReader = cmdSelectParent.ExecuteReader(CommandBehavior.CloseConnection)

        If dr.HasRows Then
            '  Response.Write(PageSize.SelectedValue)
            '  Response.End()
            rpt90097.DataSource = dr
            rpt90097.DataBind()
        End If
        con.Close()

    End Sub
    Sub GetDep()

        'If Len(arrPermissionsByDepartmentId) > 0 Then

        'End If
        Dim cmdSelect = New SqlCommand("Select departmentId, departmentName,PriorityLevel from Departments order by departmentName", con)
        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtDep)
        con.Close()
        primKeydtDep(0) = dtDep.Columns("departmentId")
        dtDep.PrimaryKey = primKeydtDep
        seldep.Items.Clear()
        '  Dim list1 As New ListItem("הכל", "0")
        '  seldep.Items.Add(list1)
        For i As Integer = 0 To dtDep.Rows.Count - 1
            Dim list As New ListItem(dtDep.Rows(i)("departmentName"), dtDep.Rows(i)("departmentId"))
            'If (Request("seldep") > 0 And Request("seldep") = dtDep.Rows(i)("departmentId")) Then
            '   Response.Write("dep=" & dep & "<BR>")
            '  Response.Write("dep=" & dtDep.Rows(i)("departmentId"))

            seldep.Items.Add(list)
        Next
        '  Response.Write("arrPermissionsByDepartmentId=" & arrPermissionsByDepartmentId & "--")
        '''   Response.End()
        If Len(arrPermissionsByDepartmentId) > 0 And arrPermissionsByDepartmentId <> "0" Then
            '  Response.Write(">0")
            '  Response.End()
            Dim parts As String() = arrPermissionsByDepartmentId.Split(",")
            For Each part As String In parts
                seldep.Items.FindByValue(part).Selected = True
            Next
            'Else
            '    Response.Write("0")
        End If

        'Response.Write("dep=" & dep)
        ' Response.End()


    End Sub

    Private Sub rptParent_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptParent.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim rptBar As Repeater
            Dim cmdSelect = New SqlCommand
            Dim parentId As Integer = e.Item.DataItem("Permission_Id")
            rptBar = e.Item.FindControl("rptBar")
            '  Dim cmdSelect = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=@parentId", conBar)
            'cmdSelect = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name, " & _
            '  " UserPermissions.is_visible, UserPermissions.User_Id, Permissions.Bar_Order FROM  Permissions RIGHT OUTER JOIN  UserPermissions ON " & _
            '  " Permissions.Bar_Id = UserPermissions.bar_id WHERE(Permissions.Parent_Id = @parentId) ORDER BY Permissions.Bar_Order", conBar)
            ' Dim cmdSelect = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=@parentId", conBar)
            cmdSelect = New SqlCommand("SELECT DISTINCT Permissions.Permission_Id, Permissions.Bar_Id, Permissions.Parent_Id, Permissions.Permission_Name,dbo.IsVisiblePermissions(Permissions.Bar_Id," & UserId & ") as is_visible, Permissions.Bar_Order FROM  Permissions  WHERE(Permissions.Parent_Id = @parentId) ORDER BY Permissions.Bar_Order", conBar)



            cmdSelect.Parameters.Add("@parentId", SqlDbType.Int).Value = CInt(parentId)

            conBar.Open()
            Dim dr As SqlDataReader = cmdSelect.ExecuteReader(CommandBehavior.CloseConnection)

            If dr.HasRows Then
                '  Response.Write(PageSize.SelectedValue)
                '  Response.End()
                rptBar.DataSource = dr
                rptBar.DataBind()

            End If

            conBar.Close()


        End If

    End Sub
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        saveData()
        Dim cScript As String

        cScript = "<script language='javascript'>self.close(); </script>"
        RegisterStartupScript("ReloadScrpt", cScript)

    End Sub
    Public Sub saveData()
        Dim conPerm As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        Dim cmd As New SqlClient.SqlCommand
        Dim i As Integer = 0
        Dim barVisible As String
        Dim barID, PermissionId As Integer
        Dim cmdSelect = New SqlCommand
        Dim cmdDelete As New SqlClient.SqlCommand("Delete FROM UserPermissions WHERE  USER_ID = @UserId;Update Users set PermissionsByDepartmentId=@seldep WHERE  USER_ID = @UserId", con)
        cmdDelete.CommandType = CommandType.Text
        cmdDelete.Parameters.Add("@UserId", SqlDbType.Int).Value = UserId
        cmdDelete.Parameters.Add("@seldep", SqlDbType.VarChar, 200).Value = Trim(Request("seldep"))
        con.Open()
        '  Try
        cmdDelete.ExecuteNonQuery()
        con.Close()
        Dim dr As SqlDataReader

        'Insert --
        ' Response.Write("ff=" & Request("seldep"))
        ' Response.Write("B=" & Request.Form("is_visible1"))
        Dim sqlstr As String
        Dim depArr As Array
        depArr = Request("seldep").Split(",")
        For i = 0 To UBound(depArr)
            '  Response.Write("i=" & i & ":" & depArr(i) & "<BR>")
            con.Open()

            cmdSelect = New SqlCommand("Select Permission_Id,bar_id from Permissions  Order by PARENT_ID, bar_Order", con)
            dr = cmdSelect.ExecuteReader()

            While dr.Read()

                barID = Trim(dr("bar_id"))
                PermissionId = Trim(dr("Permission_Id"))
                '   Response.Write("b=" & barID & "<BR>")
                If Trim(Request.Form("is_visible" & barID)) <> "" Then
                    '  Response.Write(barID & ":" & Request.Form("is_visible" & barID) & "<BR>")

                    barVisible = Request.Form("is_visible" & barID)
                End If

                '''    Response.Write("fff")
                If Trim(barVisible) = "on" Then
                    barVisible = 1
                    '  Response.Write("barVisible=" & barVisible & "<BR>")
                Else
                    barVisible = 0
                End If
                conPerm.Open()
                sqlstr = "Insert Into UserPermissions (User_Id,Department_Id,Permission_Id,bar_id,is_visible) values (" & UserId & "," & depArr(i) & "," & PermissionId & "," & PermissionId & ",'" & barVisible & "')"
                cmd = New System.Data.SqlClient.SqlCommand(sqlstr, conPerm)
                cmd.CommandType = CommandType.Text
                cmd.ExecuteNonQuery()
                conPerm.Close()
                '   Response.Write(sqlstr & "<BR>")
                '  Response.End()
            End While
            con.Close()
            dr = Nothing



            'If barID = "1" Then
            '    COMPANIES = barVisible
            'ElseIf barID = "2" Then
            '    SURVEYS = barVisible
            'ElseIf barID = "3" Then
            '    EMAILS = barVisible
            'ElseIf barID = "4" Then
            '    WORK_PRICING = barVisible
            'ElseIf barID = "5" Then
            '    PROPERTIES = barVisible
            'ElseIf barID = "6" Then
            '    TASKS = barVisible
            'ElseIf barID = "29" Then
            '    CASH_FLOW = barVisible
            'End If
            'sqlstr = "Insert Into bar_users values (" & barID & "," & OrgID & "," & USER_ID & ",'" & barVisible & "')"
            'con.executeQuery(sqlstr)
            'rs_bars.moveNext()
            'End While
            'rs_bars = Nothing
            '   i = i + 1

        Next





    End Sub


End Class

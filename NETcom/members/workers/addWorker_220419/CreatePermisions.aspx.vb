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
 
    Protected WithEvents rptParent As Repeater


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
        If IsNumeric(Request("Id")) Then
            UserId = Request("Id")
            Dim cmdSelect = New SqlCommand("SELECT  FIRSTNAME + char(32) +LASTNAME as UserName from Users where User_Id=@UserId", con)
            cmdSelect.Parameters.Add("@UserId", SqlDbType.Int).Value = UserId
            con.Open()
            UserName = cmdSelect.ExecuteScalar
            con.Close()
            If Not IsPostBack Then
                GetDep()
                GetData()
            End If

        Else
            Response.Write("No userId")
          
        End If
    End Sub
    Sub GetData()
        Dim cmdSelectParent = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=0", con)
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
    Sub GetDep()
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
        'Response.Write("dep=" & dep)
        ' Response.End()


    End Sub

    Private Sub rptParent_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.RepeaterItemEventArgs) Handles rptParent.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim rptBar As Repeater
            Dim parentId As Integer = e.Item.DataItem("Permission_Id")
            rptBar = e.Item.FindControl("rptBar")
            Dim cmdSelect = New SqlCommand("select Permission_Id,Bar_Id,Parent_Id,Permission_Name from Permissions where Parent_id=@parentId", conBar)
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
End Class

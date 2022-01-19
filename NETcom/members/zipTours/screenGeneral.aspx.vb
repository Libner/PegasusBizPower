Public Class screenGeneral
    Inherits System.Web.UI.Page
    Dim func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Protected WithEvents dtlGeneral As DataList
    Protected WithEvents btnSave As System.Web.UI.WebControls.Button



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
        If Not Page.IsPostBack Then
            cmdSelect = New SqlClient.SqlCommand("SELECT    Column_Id, Column_Name,  Column_Visible  FROM ScreenSetting  order by Column_Order", con)
            cmdSelect.CommandType = CommandType.Text

            con.Open()
            dtlGeneral.DataSource = cmdSelect.ExecuteReader()
            dtlGeneral.DataBind()
            con.Close()
        End If
    End Sub

    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click

        saveGeneral()
    End Sub
    Sub saveGeneral()
        Dim i As Integer
        Dim chk As HtmlInputCheckBox
        For i = 0 To dtlGeneral.Items.Count - 1
            chk = CType(dtlGeneral.Items(i).FindControl("chkSpec"), HtmlInputCheckBox)
                   'cmd.CommandType = CommandType.Text
            If chk.Checked Then
                con.Open()
                Dim cmd = New System.Data.SqlClient.SqlCommand("update ScreenSetting set Column_Visible=1 where column_Order=@Id", con)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@Id", SqlDbType.Int).Value = i + 1
                cmd.ExecuteNonQuery()
                con.Close()

            Else
                con.Open()
                Dim cmd = New System.Data.SqlClient.SqlCommand("update ScreenSetting set Column_Visible=0 where column_Order=@Id", con)
                cmd.CommandType = CommandType.Text
                cmd.Parameters.Add("@Id", SqlDbType.Int).Value = i + 1 'CInt(chk.Attributes("value"))
                cmd.ExecuteNonQuery()
                con.Close()
            End If
            '    Response.Write("-=" & chk.Value & "<BR>")
            ' If chk.Checked Then

            'cmd = New System.Data.SqlClient.SqlCommand("update iScreenSetting set Column_Visible=", con)
            'cmd.CommandType = CommandType.Text
            'cmd.Parameters.Add("@Id", SqlDbType.Int).Value = CInt(chk.Attributes("value"))

            'cmd.ExecuteNonQuery()
            '  End If
        Next

    End Sub

    Private Sub dtlGeneral_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataListItemEventArgs) Handles dtlGeneral.ItemDataBound

        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim chkSpec As HtmlInputCheckBox
            chkSpec = e.Item.FindControl("chkSpec")
            chkSpec.Attributes("name") = "chkSpec"

            chkSpec.Attributes.Add("value", e.Item.DataItem("Column_Id"))
            If CBool(e.Item.DataItem("Column_Visible")) Then
                chkSpec.Checked = True
           
                chkSpec.Attributes.Add("title", "לבטל שדה")
      
            Else
                chkSpec.Checked = False
                 chkSpec.Style.Add("background-color", "#d5d5d5")
                chkSpec.Attributes.Add("title", "להציג שדה")
    
            End If
        
        End If
    End Sub

End Class

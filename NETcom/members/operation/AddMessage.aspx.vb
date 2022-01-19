Imports System.Data.SqlClient
Public Class AddMessage
    Inherits System.Web.UI.Page
    Protected con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public dtSerias As New DataTable
    Dim primKeydtSerias(0) As Data.DataColumn
    Public dtSeries, dtUsers, dtGuides, dtSuppliers As New DataTable
    Dim primKeydtSeries(0), primKeydtUsers(0), primKeydtGuides(0), primKeydtSuppliers(0) As Data.DataColumn
    Protected WithEvents sSeries As System.Web.UI.HtmlControls.HtmlSelect
    Public func As New bizpower.cfunc
    Protected WithEvents sDepCode As System.Web.UI.HtmlControls.HtmlSelect
    Protected Messages_Content As System.Web.UI.HtmlControls.HtmlTextArea


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
        getSerias()
    End Sub
    Sub getSerias()
        Dim cmdSelect As New SqlClient.SqlCommand

        cmdSelect = New SqlClient.SqlCommand("select  Series_Id,Series_Name from Series ORDER BY Series_Name", con)

        con.Open()
        Dim ad As New SqlClient.SqlDataAdapter
        ad.SelectCommand = cmdSelect
        ad.Fill(dtSeries)
        con.Close()
        primKeydtSeries(0) = dtSeries.Columns("Series_Id")
        dtSeries.PrimaryKey = primKeydtSeries
        sSeries.Items.Clear()
        Dim list1 As New ListItem("הכל", "0")
        sSeries.Items.Add(list1)
        For i As Integer = 0 To dtSeries.Rows.Count - 1
            Dim list As New ListItem(dtSeries.Rows(i)("Series_Name"), dtSeries.Rows(i)("Series_Id"))
            If Request.QueryString("sSer") > 0 And Request.QueryString("sSer") = dtSeries.Rows(i)("Series_Id") Then
                list.Selected = True
            End If
            sSeries.Items.Add(list)
        Next

    End Sub

End Class

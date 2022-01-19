Public Class report_graph
    Inherits System.Web.UI.Page
    Dim func As New bizpower.cfunc
    Protected dateStart, dateEnd As String
    Protected reportUserId As Integer = 0

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
        dateStart = Trim(Request.QueryString("dateStart"))
        dateEnd = Trim(Request.QueryString("dateEnd"))
        reportUserId = func.fixNumeric(Request.QueryString("UserId"))

    End Sub

End Class

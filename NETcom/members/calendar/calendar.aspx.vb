Public Class calendar
    Inherits System.Web.UI.Page
    Protected i, j, yy As Integer
    Public dt As String
    Protected func As New include.funcs
    Protected currYear As String
    Public dayT As String
    Protected WithEvents clnEvents As calendar



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
        Session.LCID = 1037
        dt = Year(Now())
        If Not IsNumeric(Request("y")) Then
            currYear = dt
        Else
            currYear = Request("y")

        End If
    End Sub

End Class

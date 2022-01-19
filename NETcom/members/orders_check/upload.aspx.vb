Public Class Upload
    Inherits System.Web.UI.Page
    Protected strFileName As String = "" : Protected strSavePath As String = ""
    Protected str_mappath As String = "" : Protected table As String = "" : Protected object_name As String = ""

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
        If IsNothing(Request.Cookies("bizpegasus")) Then
            Response.Redirect("/")
        End If

        If Not IsNumeric(Request.Cookies("bizpegasus")("UserId")) Then
            Response.Redirect("/")
        End If

        table = Trim(Request.QueryString("table"))
        Select Case LCase(table)
            Case "gilboa" : object_name = "דוח גלבוע"
        End Select

        str_mappath = Application("VirDir") & "/download/import"

        If Request.Files.Count > 0 Then
            If Request.Files(0).ContentLength > 0 Then
                strFileName = Trim(table) & ".xls"
                strSavePath = Server.MapPath(str_mappath & "/" & strFileName)
                Request.Files(0).SaveAs(strSavePath)
                Response.Redirect("Read_Excel.aspx?table=" & table)
            End If
        End If

    End Sub

End Class

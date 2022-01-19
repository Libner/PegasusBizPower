Imports System.Net
Imports System.Net.Sockets

Public Class testEmail
    Inherits System.Web.UI.Page

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
        Dim address = "olga@eltam.com"

        Dim host() As String = (address.Split("@"))
        Dim hostname As String = host(1)

        Dim IPhst As IPHostEntry = Dns.Resolve(hostname)
        Dim endPt As New IPEndPoint(IPhst.AddressList(0), 25)

        Dim s As Socket = New Socket(endPt.AddressFamily, SocketType.Stream, ProtocolType.Tcp)
        Try
            s.Connect(endPt)
        Catch ex As Exception
            Response.Write(ex.ToString)
        End Try



    End Sub

End Class

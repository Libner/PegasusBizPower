Imports System.Data.SqlClient
Imports System.Xml
Imports System.Net
Imports System.Web.Services.Protocols
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.COMException

Public Class CheckPaxFile
    Inherits System.Web.UI.Page
    Public quest_id, appid, appDocket As String

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
        quest_id = Request("quest_id")
        appid = Request("appid")
        appDocket = Request("appDocket")
        Response.Write(quest_id & ":" & appid & ":" & appDocket)
    End Sub

End Class

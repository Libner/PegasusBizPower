Imports System.Web
Imports System.Web.SessionState

Public Class Global
    Inherits System.Web.HttpApplication

#Region " Component Designer Generated Code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
    End Sub

#End Region

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
 
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the session is started

    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires at the beginning of each request     
        ' Fires when the application is started
        HttpContext.Current.Application.Lock()
        HttpContext.Current.Application("VirDir") = HttpContext.Current.Request.ApplicationPath
        If Trim(HttpContext.Current.Application("VirDir")) = "/" Then
            HttpContext.Current.Application("VirDir") = ""
        End If

        HttpContext.Current.Application.UnLock()
    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires upon attempting to authenticate the use
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when an error occurs
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the session ends
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the application ends
    End Sub

    Public Overrides Function GetVaryByCustomString(ByVal context As HttpContext, ByVal arg As String) As String
      
        If (arg = "xmlChanged;cookieCode") Then
            Dim xmlChanged As String = System.IO.File.GetLastWriteTime(Server.MapPath(context.Application("VirDir") & _
            "/download/xml_forms/bizpower_forms.xml")).ToShortTimeString()

            Dim UserId As String = ""
            If Not IsNothing(context.Request.Cookies("bizpegasus")) Then
                UserId = context.Request.Cookies("bizpegasus")("UserId")
            End If

            Return xmlChanged & "|" & UserId
        End If

        Return MyBase.GetVaryByCustomString(context, arg)
    End Function

End Class

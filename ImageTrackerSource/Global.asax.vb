Imports System.Web
Imports System.Web.SessionState
Imports System.Text

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
        ' Fires when the application is started
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the session is started
    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires at the beginning of each request
    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires upon attempting to authenticate the use
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when an error occurs
        Dim Msg
        Msg = Server.CreateObject("CDO.Message")
        Msg.BodyPart.Charset = "windows-1255"
        Msg.From = "olga@eltam.com"
        Msg.MimeFormatted = True
        Msg.Subject = "Error in the BizPower"
        Msg.To = "bizpower.cyber@gmail.com"
        Msg.HTMLBody = BuildMessage()
        Msg.HTMLBodyPart.ContentTransferEncoding = "quoted-printable"
        Msg.Send()
        Msg = Nothing
    End Sub

    Function BuildMessage() As String

        Dim strMessage As New StringBuilder

        strMessage.Append("<style type=""text/css"">")
        strMessage.Append("<!--")
        strMessage.Append(".basix {")
        strMessage.Append("font-family: Verdana, Arial, Helvetica, sans-serif;")
        strMessage.Append("font-size: 12px;")
        strMessage.Append("}")
        strMessage.Append(".header1 {")
        strMessage.Append("font-family: Verdana, Arial, Helvetica, sans-serif;")
        strMessage.Append("font-size: 12px;")
        strMessage.Append("font-weight: bold;")
        strMessage.Append("color: #000099;")
        strMessage.Append("}")
        strMessage.Append(".tlbbkground1 {")
        strMessage.Append("background-color: #000099;")
        strMessage.Append("}")
        strMessage.Append("-->")
        strMessage.Append("</style>")

        strMessage.Append("<table width=""85%"" border=""0"" align=""center"" cellpadding=""5"" cellspacing=""1"" class=""tlbbkground1"">")
        strMessage.Append("<tr bgcolor=""#eeeeee"">")
        strMessage.Append("<td colspan=""2"" class=""header1"">Page Error</td>")
        strMessage.Append("</tr>")
        strMessage.Append("<tr>")
        strMessage.Append("<td width=""100"" align=""right"" bgcolor=""#eeeeee"" class=""header1"" nowrap>IP Address</td>")
        strMessage.Append("<td bgcolor=""#FFFFFF"" class=""basix"">" & Request.ServerVariables("REMOTE_ADDR") & "</td>")
        strMessage.Append("</tr>")
        strMessage.Append("<tr>")
        strMessage.Append("<td width=""100"" align=""right"" bgcolor=""#eeeeee"" class=""header1"" nowrap>User Agent</td>")
        strMessage.Append("<td bgcolor=""#FFFFFF"" class=""basix"">" & Request.ServerVariables("HTTP_USER_AGENT") & "</td>")
        strMessage.Append("</tr>")
        strMessage.Append("<tr>")
        strMessage.Append("<td width=""100"" align=""right"" bgcolor=""#eeeeee"" class=""header1"" nowrap>Page</td>")
        strMessage.Append("<td bgcolor=""#FFFFFF"" class=""basix"">" & Request.Url.AbsoluteUri & "</td>")
        strMessage.Append("</tr>")
        strMessage.Append("<tr>")
        strMessage.Append("<td width=""100"" align=""right"" bgcolor=""#eeeeee"" class=""header1"" nowrap>Time</td>")
        strMessage.Append("<td bgcolor=""#FFFFFF"" class=""basix"">" & System.DateTime.Now & " EST</td>")
        strMessage.Append("</tr>")
        strMessage.Append("<tr>")
        strMessage.Append("<td width=""100"" align=""right"" bgcolor=""#eeeeee"" class=""header1"" nowrap>Details</td>")
        strMessage.Append("<td bgcolor=""#FFFFFF"" class=""basix"">" & Server.GetLastError().InnerException.ToString() & "</td>")
        strMessage.Append("</tr>")
        strMessage.Append("</table>")
        Return strMessage.ToString
    End Function

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the session ends
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the application ends
    End Sub

End Class

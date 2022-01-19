Public Class test2
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
        Dim Msg As New System.Web.Mail.MailMessage
        Dim EmailContent As String

        EmailContent = "Please enter Your Pegasus account to view the customers feedbacks regarding group "

        Msg.BodyFormat = System.Web.Mail.MailFormat.Html
        Msg.BodyEncoding = System.Text.Encoding.UTF8
        Msg.From = "info@pegasusisrael.co.il"
        Msg.Subject = "test Feedbacks"
        Msg.To = "faina@cyberserve.co.il"
              Msg.Body = EmailContent.ToString()

        Try
            System.Web.Mail.SmtpMail.SmtpServer = ConfigurationSettings.AppSettings.Item("smtp_server")
            System.Web.Mail.SmtpMail.Send(Msg)

            '    Response.Write("ok")
        Catch ex As Exception

        End Try
        Msg = Nothing
    End Sub

End Class

Public Class SEndGuidePdfReportMail
    Inherits System.Web.UI.Page
    Protected gId, selUserID As Integer
    Dim func As New bizpower.cfunc

    Protected gyear, Guide_Name, Guide_Email, selUserName, EMAIL As String
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Public WithEvents btnSend As Button

    Public rs_Guide, rs_Users As SqlClient.SqlDataReader


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
        '  btnSend.Attributes.Add("onclick", "return CheckFields(1);")
        gId = Request("gId")
        gyear = Request("gyear")
        Dim sqlstrGuide As New SqlClient.SqlCommand("SELECT Guide_Id, (Guide_FName + ' - ' + Guide_LName) as Guide_Name,Guide_Email  FROM Guides  where Guide_Id=" & gId, conPegasus)
        sqlstrGuide.CommandType = CommandType.Text
        conPegasus.Open()
        rs_Guide = sqlstrGuide.ExecuteReader()
        While rs_Guide.Read()
            Guide_Name = rs_Guide("Guide_Name")
            Guide_Email = rs_Guide("Guide_Email")
        End While
        conPegasus.Close()
        Dim sqlstrUser As New SqlClient.SqlCommand("SELECT User_ID,FIRSTNAME + ' ' + LASTNAME as UserName,EMAIL FROM Users WHERE User_id=" & Request.Cookies("bizpegasus")("UserId"), con)
        sqlstrUser.CommandType = CommandType.Text
        con.Open()
        rs_Users = sqlstrUser.ExecuteReader()
        While rs_Users.Read()
            selUserID = rs_Users("User_ID")
            selUserName = rs_Users("UserName")
            EMAIL = rs_Users("EMAIL")
        End While
        con.Close()
        If IsPostBack Then
            sendMail()
        End If
    End Sub

    Private Sub sendMail()
        Dim ReplyText, SubjectText, sender_name, sendermail As String
        ReplyText = Trim(Request.Form("ReplyText"))
        SubjectText = Trim(Request.Form("SubjectText"))
        sender_name = Trim(Request.Form("sender_name"))
        sendermail = Request.Form("sendermail") 'mail
        selUserID = Request.Cookies("bizpegasus")("UserId")
        Dim strBody = "<html><head><title>BizPower</title><meta http-equiv=Content-Type content=""text/html; charset=windows-1255"">" & Chr(10) & Chr(13) & _
  "<link href= /netcom/IE4.css rel=STYLESHEET type=text/css></head><body>" & Chr(10) & Chr(13) & _
  "<table border=0 width=100% cellspacing=0 cellpadding=0 align=right>"
        strBody = strBody & "<tr><td width=100% align=right valign=top height=20 nowrap></td></tr>" & vbCrLf
        If Len(ReplyText) > 0 Then
            strBody = strBody & "<tr><td align=""right"" width=100% ><span  dir=""rtl"">" & func.breaks(Trim(ReplyText)) & "</span></td></tr>"
        End If
        strBody = strBody & "<tr><td height=20></td></tr>"

        strBody = strBody & "<tr><td height=30 nowrap colspan=2></td></tr>" & Chr(10) & Chr(13) & _
     "</table></td></tr></table>"


        '   wFile = "https://pegasusisrael.co.il/biz_form/" + fname + "?guide_id=" + document.getElementById("guide_id").value + "&y=" + document.getElementById("currentYear").value



        Dim Msg As New System.Web.Mail.MailMessage
        Msg.BodyFormat = System.Web.Mail.MailFormat.Html
        Msg.BodyEncoding = System.Text.Encoding.UTF8
        Msg.From = sender_name & "<" & sendermail & ">"
        Msg.Subject = SubjectText
        Msg.To = "faina@cyberserve.co.il"  'Request.Form("Guide_Email")
        Msg.Body = strBody.ToString()
        System.Web.Mail.SmtpMail.Send(Msg)
        Msg = Nothing




    End Sub
End Class

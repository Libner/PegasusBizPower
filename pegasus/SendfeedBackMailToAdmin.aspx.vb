Public Class SendfeedBackMailToAdmin
    Inherits System.Web.UI.Page
    Public dr_m As SqlClient.SqlDataReader
    Public ContactId, MemberId, EmailStatusId, companyID, UserID As Integer
    Public EmailTo, Email_phone, Emailtype_id, orgName, DateBegin, TourTitle, GuideName, FeedBack_Url, FeedBack_content As String
    Protected DepartureId As String = 0
    Protected TourGrade As Double

    Protected ImportanceId As Integer
    Public func As New bizpower.cfunc
    Protected Email_content, FromEmail As String
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))

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

        Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
        companyID = 12681
        UserID = 1028    'מנהל אתר אינטרנט

        FromEmail = "info@pegasusisrael.co.il"
    
        Dim cmdSelect As New SqlClient.SqlCommand("SET DATEFORMAT dmy;SELECT Tours_Departures.Departure_Id,Departure_Code,Tours_Departures.Tour_Id,Departure_Date,Departure_Date_End ,Tour_Name=case when isnull(Tour_PageTitle,'')='' then Tour_Name else Tour_PageTitle end FROM   dbo.Tours_Departures " & _
        " left join Tours on Tours.Tour_id=Tours_Departures.Tour_Id " & _
        " where (datediff(dd,Departure_Date_End,GetDate())=10) ", conPegasus)
        cmdSelect.CommandType = CommandType.Text
        conPegasus.Open()
        dr_m = cmdSelect.ExecuteReader()
        Dim urlF
        While dr_m.Read

         
            If Not dr_m("Departure_Id") Is DBNull.Value Then
                DepartureId = dr_m("Departure_Id")
            End If


            urlF = "http://pegasus.bizpower.co.il/netcom/members/feedBack/default_gradesByDepId.asp?depId=" & DepartureId
            TourGrade = Replace(func.TourGrade(DepartureId), "%", "")
            Select Case TourGrade
                Case 0 To 60
                    ImportanceId = 4
                    EmailTo = func.EmailToByImportance(ImportanceId)

                Case 61 To 75
                    ImportanceId = 3
                    EmailTo = func.EmailToByImportance(ImportanceId)
                Case 76 To 80
                    ImportanceId = 2
                    EmailTo = func.EmailToByImportance(ImportanceId)
                Case 81 To 90
                    ImportanceId = 1
                    EmailTo = func.EmailToByImportance(ImportanceId)
                Case Else
                    ImportanceId = 0
                    EmailTo = ""

            End Select
            Dim strBody As New System.Text.StringBuilder
            strBody.Append("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"" ""http://www.w3.org/TR/html4/loose.dtd"">" & vbCrLf)
            strBody.Append("<html><head><title></title>" & vbCrLf)
            strBody.Append("<style type=""text/css""> .page_title{" & vbCrLf)
            strBody.Append(" font-family:Arial;" & vbCrLf)
            strBody.Append(" FONT-WEIGHT: bolder;" & vbCrLf)
            strBody.Append(" font-size:12pt;" & vbCrLf)
            strBody.Append(" color:#E60009; " & vbCrLf)
            strBody.Append(" background-color:#E0E0E0;" & vbCrLf)
            strBody.Append(" padding-left:10px;" & vbCrLf)
            strBody.Append(" padding-right:10px;" & vbCrLf)
            strBody.Append(" line-height: 18px;" & vbCrLf)
            strBody.Append(" height: 26px;}" & vbCrLf)
            strBody.Append(" BODY" & vbCrLf)
            strBody.Append(" {" & vbCrLf)
            strBody.Append("  margin: 0px;" & vbCrLf)
            strBody.Append("  font-family:Arial;" & vbCrLf)
            strBody.Append("  COLOR: #666666;" & vbCrLf)
            strBody.Append(" }" & vbCrLf)
            strBody.Append(" TD" & vbCrLf)
            strBody.Append(" {" & vbCrLf)
            strBody.Append("  text-align: right;" & vbCrLf)
            ' strBody.Append("  direction: rtl;" & vbCrLf)
            strBody.Append(" }" & vbCrLf)
            strBody.Append("</style></head><body>" & vbCrLf)
            strBody.Append("<table cellpadding=3 cellspacing=1 border=0 width=550 align=center>" & vbCrLf)
            strBody.Append("<tr><td colspan=""2"">:שים לב לציוני הטיול </td></tr>" & vbCrLf)
            strBody.Append("<tr><td colspan=""2""><b>" & dr_m("Tour_Name") & "</b></td></tr>" & vbCrLf)
            strBody.Append("<tr><td colspan=""2"">" & func.TourGrade(DepartureId) & " טיול זה חזר בציון </td></tr>" & vbCrLf)
            strBody.Append("<tr><td colspan=""2"">.אנא פעל לפי הנוהלים בהקדם</td></tr>" & vbCrLf)
            strBody.Append("<tr><td colspan=""2""><a href=" & urlF & ">לחץ כאן" & "</a></td></tr>" & vbCrLf)
            strBody.Append("</table></body></html>")


            '  FeedBack_content = FeedBack_content & urlF
            ' FeedBack_content = FeedBack_content & "<BR>" & EmailTo & " -" & TourGrade & "-" & ImportanceId
            ' FeedBack_content = FeedBack_content '& "<BR>" & EmailTo & " -" & TourGrade & "-" & ImportanceId
            '   Email_content = FeedBack_content
            '  Response.Write(FeedBack_content & "<BR>")
            If Trim(EmailTo) <> "" And ImportanceId > 0 Then
                Dim Msg As New System.Web.Mail.MailMessage
                Msg.BodyFormat = System.Web.Mail.MailFormat.Html
                Msg.BodyEncoding = System.Text.Encoding.UTF8
                Msg.From = "info@pegasusisrael.co.il"
                Msg.Subject = "  משוב לטיול - " & dr_m("Tour_Name")
                Msg.To = EmailTo '"faina@cyberserve.co.il" 'EmailTo
                Msg.Body = strBody.ToString()
                System.Web.Mail.SmtpMail.Send(Msg)
                Msg = Nothing
            End If




        End While
        dr_m.Close()
        cmdSelect.Dispose()
        conPegasus.Close()

        'Response.Write("send mail- end ")

    End Sub

End Class

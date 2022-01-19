Imports System.Data.SqlClient
Imports System.Web.Mail
Public Class SendMailRishumFromMitanen
    Inherits System.Web.UI.Page
    Public func As New bizpower.cfunc
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Dim cmdSelect As New SqlClient.SqlCommand
    Dim BizpowerDBName As String = ConfigurationSettings.AppSettings("BizpowerDBName")
    Dim SiteDBName As String = ConfigurationSettings.AppSettings("PegasusSiteDBName")
    Dim isSuccess As Boolean = False
    Public appealsCount As Integer = 0 ' total count of selected appeals
    Public emailSentCount As Integer = 0 ' total count of selected appeals
    Public sentCount As Integer = 0 ' count of appeals for which the mail was sent

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

        Dim userId, orgID, lang_id, DepartmentId, companyID As String

        Dim appealsId As String ' appaels ID for which to send mail
        Dim apealsIsMarketingEmailSend As String ' appaels ID for which the mail was sent 

        userId = Trim(Request.Cookies("bizpegasus")("UserId"))
        orgID = Trim(Request.Cookies("bizpegasus")("OrgID"))
        lang_id = Trim(Request.Cookies("bizpegasus")("LANGID"))

        appealsId = Request.QueryString("appealsId") 'get appaels ID

        If appealsId <> "" Then
            Dim appList() As String
            appList = appealsId.Split(",")

            If appList.Length > 0 Then
                Dim strBody As String ' text to sending ('XXXXXXXXXX - replace to link according to tour)
                strBody = ""

                strBody = strBody & (" <!DOCTYPE html>")
                strBody = strBody & ("<html><head><title></title>" & vbCrLf)
                strBody = strBody & ("<style>")
                strBody = strBody & ("td.content_cl")
                strBody = strBody & ("{")
                strBody = strBody & (vbTab & "FONT-WEIGHT: normal;")
                strBody = strBody & (vbTab & "FONT-SIZE: 11pt;")
                strBody = strBody & (vbTab & "COLOR: #193B6E;")
                strBody = strBody & (vbTab & "FONT-FAMILY: Arial;")
                strBody = strBody & ("}")
                strBody = strBody & ("td.content_b")
                strBody = strBody & ("{")
                strBody = strBody & (vbTab & "FONT-WEIGHT: bold;")
                strBody = strBody & (vbTab & "FONT-SIZE: 11pt;")
                strBody = strBody & (vbTab & "COLOR: #193B6E;")
                strBody = strBody & (vbTab & "FONT-FAMILY: Arial;")
                strBody = strBody & ("}")
                strBody = strBody & ("td.content_Title")
                strBody = strBody & ("{")
                strBody = strBody & (vbTab & "FONT-WEIGHT: bold;")
                strBody = strBody & (vbTab & "FONT-SIZE: 18pt;")
                strBody = strBody & (vbTab & "COLOR: #193B6E;")
                strBody = strBody & (vbTab & "FONT-FAMILY: Arial;")
                strBody = strBody & ("}")
                strBody = strBody & (".button")
                strBody = strBody & ("{")
                strBody = strBody & (vbTab & "display:block;")
                strBody = strBody & (vbTab & "width:auto;")
                strBody = strBody & (vbTab & "height: auto;")
                strBody = strBody & (vbTab & "margin: 0px;")
                strBody = strBody & (vbTab & "padding: 10px;")
                strBody = strBody & (vbTab & "border-radius:   8px;")
                strBody = strBody & (vbTab & "moz-border - radius:   8px;")
                strBody = strBody & (vbTab & "khtml-border - radius:   8px;")
                strBody = strBody & (vbTab & "o-border - radius:   8px;")
                strBody = strBody & (vbTab & "webkit-border - radius:   8px;")
                strBody = strBody & (vbTab & "ms-border - radius:   8px;  ")
                strBody = strBody & (vbTab & "background-color:  #93bd2a;")
                strBody = strBody & (vbTab & "color: #ffffff;")
                strBody = strBody & (vbTab & "Font-Size:   20pt;")
                strBody = strBody & (vbTab & "Font-weight:  700;")
                strBody = strBody & (vbTab & "text-align:   center;")
                strBody = strBody & (vbTab & "text-decoration:none;")
                strBody = strBody & ("}")
                strBody = strBody & ("</style></head><body>" & vbCrLf)
                strBody = strBody & ("<table cellpadding=0 cellspacing=5 align='center' border='0' align=right>")
                strBody = strBody & ("<tr>")
                strBody = strBody & ("<td align='right' dir='rtl' class='content_b' nowrap width='100%'></td>")
                strBody = strBody & ("</tr>")
                strBody = strBody & ("<tr>")
                strBody = strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'>מטיילים יקרים</td>")
                strBody = strBody & ("</tr>")
                strBody = strBody & ("<tr>")
                strBody = strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'>תודה על התעניינותכם לטיול לCCCCCCCCCC.<bR></td>")
                strBody = strBody & ("</tr>")

                strBody = strBody & ("<tr>")
                strBody = strBody & ("<td bgcolor='#f5f5f5' align='right'>")
                strBody = strBody & ("<table cellspacing='0' cellpadding='10'>")
                strBody = strBody & ("<tr>")
                strBody = strBody & ("<td align='right' dir='rtl' class='content_Title' nowrap >")
                strBody = strBody & ("<table cellspacing='0' cellpadding='10'>")
                strBody = strBody & ("<tr>")
                strBody = strBody & ("<td bgcolor='#93bd2a'>")
                strBody = strBody & ("<a href='XXXXXXXXXX'  class='button'>לחץ כאן</a>")
                strBody = strBody & ("</td>")
                strBody = strBody & ("</tr>")
                strBody = strBody & ("</table>")
                strBody = strBody & ("</td>")
                strBody = strBody & ("<td align='right' dir='rtl' class='content_Title' nowrap >קישור להרשמה אינטרנטית לטיול לCCCCCCCCCC </td>")
                strBody = strBody & ("</tr>")
                strBody = strBody & ("</table>")
                strBody = strBody & ("</td>")
                strBody = strBody & ("</tr>")
                strBody = strBody & ("<tr>")
                strBody = strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'><span style='font-size:12pt;font-weight:bold;color:red'>הרישום לטיול זה דרך האינטרנט הוא אישי: ""אין להעבירו לאדם אחר כיוון שעלולות להיגרם תקלות טכניות ולפגום ברישום""</span><br></td>")
                strBody = strBody & ("</tr>")

                strBody = strBody & ("<tr>")
                strBody = strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'>אנו ממליצים בחום להירשם ולהבטיח מקום בהקדם.</td>")
                strBody = strBody & ("</tr>")

                strBody = strBody & ("<tr>")
                strBody = strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'>שימו לב כי המחיר נשמר רק לאחר אישור ההרשמה המלווה בתשלום. <br>הירשמו מוקדם כיוון שבמידה ומחיר הטיול ירד לאחר רישום לטיול, מדיניות פגסוס היא תמיד לשמור על המטייל <BR>שנרשם מוקדם ונוריד את המחיר למחיר המעודכן.<br></td>")
                strBody = strBody & ("</tr>")

                strBody = strBody & ("<tr>")
                strBody = strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'>אנו ממליצים להקיף בעיגול כי הנכם מעוניינים לשמוע הצעת ביטוח רפואי המכסה דמי ביטול מיד עם הרשמתכם.<br></td>")
                strBody = strBody & ("</tr>")


                strBody = strBody & ("<tr>")
                strBody = strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'>לפקס 03-5436060 או למייל <a href='mailto:pegasus@pegasusisrael.co.il'>pegasus@pegasusisrael.co.il</a></td>")
                strBody = strBody & ("</tr>")


                strBody = strBody & ("<tr>")
                strBody = strBody & ("<td align='right' dir='rtl' class='content_b' nowrap width='100%'></td>")
                strBody = strBody & ("</tr>")
                strBody = strBody & ("<tr>")
                strBody = strBody & ("<td align='right' dir='rtl' class='content_cl' nowrap width='100%'>להתחלת תהליך רישום   <a href='XXXXXXXXXX' >לחץ כאן</a><br><br></td>")
                strBody = strBody & ("</tr>")

                strBody = strBody & ("<tr>")
                strBody = strBody & ("<td align='right' dir='rtl' class='content_b' nowrap width='100%'></td>")
                strBody = strBody & ("</tr>")
                strBody = strBody & ("<tr><td align='right' class='content_cl'  valign=top>טיול מהנה ומפלא</td></tr>")


                strBody = strBody & ("<tr><td align='right' class='content_cl'  valign=top>צוות מכירות פגסוס</td></tr>")

                strBody = strBody & ("</table></body></html>")


                Dim ContactId, CountryId, TourId, SeriasId As Integer
                Dim Series_Name, Country_Name As String
                Dim emailContact As String
                Dim appealId As Integer = 0
                Dim status As String

                If func.fixNumeric(userId) > 0 Then

                    'Response.Write("<br>Logs=" & Logs)
                    For i As Integer = 0 To appList.Length - 1
                        appealsCount = appealsCount + 1

                        appealId = func.fixNumeric(appList(i))
                        Country_Name = ""

                        'Response.Write("<br>LogId=" & LogId)
                        If appealId > 0 Then
                            'seng new appeal
                            cmdSelect = New SqlClient.SqlCommand("Select appeals.Contact_Id,appeal_CountryId,Tour_Id,email,Country_Name from appeals " & _
                            " left join Contacts on appeals.Contact_id=contacts.Contact_Id " & _
                            " left join Countries on appeals.appeal_CountryId=Countries.Country_Id " & _
                            " WHERE  len(email)>0 and appeal_id =@appealId", con)
                            cmdSelect.CommandType = CommandType.Text
                            cmdSelect.Parameters.Add("@appealId", SqlDbType.Int).Value = appealId
                            con.Open()
                            Dim rd As SqlDataReader
                            rd = cmdSelect.ExecuteReader
                            If rd.Read Then
                                ContactId = func.fixNumeric(rd("Contact_Id"))
                                CountryId = func.fixNumeric(rd("appeal_CountryId"))
                                TourId = func.fixNumeric(rd("Tour_Id"))
                                emailContact = func.dbNullFix(rd("email"))
                                Country_Name = func.dbNullFix(rd("Country_Name"))
                            End If
                            con.Close()

                            If ContactId > 0 And emailContact <> "" Then
                                'send
                                'set link to galor - replace XXXXXXXXXX to real link
                                Dim strAppBody As String = ""
                                Dim linkGalor As String = ""
                                linkGalor = ConfigurationSettings.AppSettings.Item("PegasusUrl") & "/tours/tourRegistrationGalorContact.aspx?appId=" & appealId
                                strAppBody = strBody
                                strAppBody = Replace(strBody, "XXXXXXXXXX", linkGalor)
                                strAppBody = Replace(strAppBody, "CCCCCCCCCC", Country_Name)

                                If Trim(emailContact) <> "" Then
                                    Dim isSent As Boolean = False
                                    Dim Msg As New System.Web.Mail.MailMessage
                                    Msg.BodyFormat = System.Web.Mail.MailFormat.Html
                                    Msg.BodyEncoding = System.Text.Encoding.UTF8
                                    Msg.From = "info@pegasusisrael.co.il"
                                    Msg.Subject = "אתר פגסוס - טופס רישום לטיול"
                                    Msg.To = emailContact
                                    Msg.Body = strAppBody.ToString()
                                    Try
                                        System.Web.Mail.SmtpMail.Send(Msg)
                                        isSent = True
                                    Catch ex As Exception
                                        Response.Write("<br>System.Web.Mail.SmtpMail -" & emailContact & " - " & Err.Description)
                                    End Try
                                    Msg = Nothing

                                    If isSent Then

                                        If apealsIsMarketingEmailSend <> "" Then
                                            apealsIsMarketingEmailSend = apealsIsMarketingEmailSend & ","
                                        End If
                                        apealsIsMarketingEmailSend = apealsIsMarketingEmailSend & appealId
                                        emailSentCount = emailSentCount + 1
                                        '===========================
                                        'Add to marketing mailing SQL tqble
                                        Dim sqlInsMailing As New SqlClient.SqlCommand("INSERT INTO dbo.MarketingMailing " & _
             "  (CONTACT_ID, USER_ID, APPEAL_ID, Country_ID, Tour_ID, MailingType_ID, Recepient_Email, Subject_Email, Content_Email, DATE_SEND, DATE_OPENED, IS_OPENED)" & _
             " VALUES (@CONTACT_ID, @USER_ID, @APPEAL_ID, @Country_ID, @Tour_ID, @MailingType_ID, @Recepient_Email, @Subject_Email, @Content_Email, getDate(), null, 0)", con)
                                        sqlInsMailing.CommandType = CommandType.Text
                                        sqlInsMailing.Parameters.Add("@CONTACT_ID", SqlDbType.Int).Value = ContactId
                                        sqlInsMailing.Parameters.Add("@USER_ID", SqlDbType.Int).Value = userId
                                        sqlInsMailing.Parameters.Add("@APPEAL_ID", SqlDbType.Int).Value = appealId
                                        sqlInsMailing.Parameters.Add("@Country_ID", SqlDbType.Int).Value = CountryId
                                        If TourId > 0 Then
                                            sqlInsMailing.Parameters.Add("@Tour_ID", SqlDbType.Int).Value = TourId
                                        Else
                                            sqlInsMailing.Parameters.Add("@Tour_ID", SqlDbType.Int).Value = DBNull.Value
                                        End If
                                        sqlInsMailing.Parameters.Add("@MailingType_ID", SqlDbType.Int).Value = 1 ' by default - from tofes mitanen (16504)
                                        sqlInsMailing.Parameters.Add("@Recepient_Email", SqlDbType.VarChar, 100).Value = emailContact
                                        sqlInsMailing.Parameters.Add("@Subject_Email", SqlDbType.VarChar, 100).Value = "אתר פגסוס - טופס רישום לטיול"
                                        sqlInsMailing.Parameters.Add("@Content_Email", SqlDbType.NText).Value = strAppBody
                                        con.Open()
                                        Try
                                            sqlInsMailing.ExecuteNonQuery()
                                            con.Close()
                                        Catch ex As Exception
                                            Response.Write("<br>MarketingMailing -" & emailContact & " - " & Err.Description)
                                        Finally
                                            con.Close()
                                        End Try
                                        '===========================
                                        '===========================
                                        'update email sending counter in appeals
                                        Dim sqlUpdMailing As New SqlClient.SqlCommand("UPDATE APPEALS SET IsMarketingEmailSend=IsMarketingEmailSend + 1 " & _
             "  WHERE APPEAL_ID=@APPEAL_ID", con)
                                        sqlUpdMailing.CommandType = CommandType.Text
                                        sqlUpdMailing.Parameters.Add("@APPEAL_ID", SqlDbType.Int).Value = appealId
                                        con.Open()
                                        Try
                                            sqlUpdMailing.ExecuteNonQuery()
                                            con.Close()
                                        Catch ex As Exception
                                            Response.Write("<br>MarketingMailing -UPDATE IsMarketingEmailSend  - " & Err.Description)
                                        Finally
                                            con.Close()
                                        End Try
                                        '===========================
                                    End If
                                End If

                            End If


                        End If
                    Next
                End If

            End If
        End If
    End Sub

End Class

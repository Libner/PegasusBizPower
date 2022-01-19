Public Class SendfeedBack_FormByMail_IsTravelers
    Inherits System.Web.UI.Page
    Public dr_m As SqlClient.SqlDataReader
    Public ContactId, TravelerId, EmailStatusId, companyID, UserID As Integer
    Public Email, Email_phone, Emailtype_id, orgName, DateBegin, TourTitle, GuideName, FeedBack_Url, FeedBack_content As String
    Protected DepartureId As Integer = 0
    Protected Email_content, FromEmail, TravelerNAME As String
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Dim conPegasus As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionStringS"))
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Public BizpowerDBName As String = ConfigurationSettings.AppSettings("BizpowerDBName")
    Public PegasusSiteDBName As String = ConfigurationSettings.AppSettings("PegasusSiteDBName")


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
        Emailtype_id = 8   'משוב טיול
        Emailtype_id = 88   'משוב טיול למטייל

        'datediff(dd,Departure_Date_End,GetDate())>-3
        Dim cmdSelect As New SqlClient.SqlCommand("SET DATEFORMAT dmy;SELECT  T.FirstName+' '+T.LastName as _Traveler_NAME , Tours_Departures.Departure_Id,Departure_Code,Tours_Departures.Tour_Id,Departure_Date,Departure_Date_End ," & _
          " TT.Traveler_Id,TT.Email,TT.Contact_Id FROM   dbo.Tours_Departures left join  " & BizpowerDBName & ".dbo.Tour_Travelers TT  on TT.Departure_Id=Tours_Departures.Departure_Id " & _
          " left join  " & BizpowerDBName & ".dbo.Travelers T  on T.Traveler_Id=TT.Traveler_Id " & _
          " where(IsTravelersTour = 1 And isCanceled = 0)" & _
          " and datediff(dd,Departure_Date_End,GetDate())=1 and (TT.Email<>'' or TT.Email<>null) and  " & BizpowerDBName & ".dbo.GetIsBitulim(Tours_Departures.Departure_Id,Contact_Id)=0 ", conPegasus)
        cmdSelect.CommandType = CommandType.Text
        conPegasus.Open()
        dr_m = cmdSelect.ExecuteReader()
        While dr_m.Read

          
            If Not dr_m("Email") Is DBNull.Value Then
                Email = dr_m("Email")
                'Email = "furfaina@gmail.com"
                ' Email = "anna@cyberserve.co.il"

            End If
            If Not dr_m("Contact_Id") Is DBNull.Value Then
                ContactId = dr_m("Contact_Id")
            End If

            If Not dr_m("Departure_Id") Is DBNull.Value Then
                DepartureId = dr_m("Departure_Id")
            End If
            If Not dr_m("Traveler_Id") Is DBNull.Value Then
                TravelerId = dr_m("Traveler_Id")
            End If
            If Not dr_m("_Traveler_NAME") Is DBNull.Value Then
                TravelerNAME = dr_m("_Traveler_NAME")
            End If




            FeedBack_Url = ConfigurationSettings.AppSettings.Item("PegasusFeedbackUrl") & "/feedback/faqIsTraveler.aspx?Email=1&DepartureId=" & DepartureId & "&ContactId=" & ContactId & "&TravelerId=" & TravelerId
            FeedBack_Url = "url=" & Server.UrlEncode(FeedBack_Url)
            Dim xmlhttpShort As Object
            Dim sendUrlShort, strResponseShort, urlF As String
            sendUrlShort = "http://tinyurl.com/api-create.php"
            'sendUrlShort = "http://tiny-url.info/api/v1/create"
            xmlhttpShort = CreateObject("MSXML2.ServerXMLHTTP")
            xmlhttpShort.open("POST", sendUrlShort, False)

            xmlhttpShort.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
            xmlhttpShort.send(FeedBack_Url & "&charset=iso-8859-8")
            strResponseShort = xmlhttpShort.responseText
            urlF = strResponseShort

            urlF = strResponseShort

            xmlhttpShort = Nothing
            FeedBack_content = TravelerNAME & ","
            FeedBack_content = FeedBack_content & "<BR>"
            FeedBack_content = FeedBack_content & "<BR> ברוך שובך ארצה <BR>נשמח לקבלת משוב על הטיול בלחיצה על הקישור <BR>"
            FeedBack_content = FeedBack_content & urlF
            FeedBack_content = FeedBack_content & "<BR>" & " תודה, פגסוס "
            Email_content = FeedBack_content

            If Trim(Email) <> "" Then
                Dim Msg As New System.Web.Mail.MailMessage
                Msg.BodyFormat = System.Web.Mail.MailFormat.Html
                Msg.BodyEncoding = System.Text.Encoding.UTF8
                Msg.From = "info@pegasusisrael.co.il"
                Msg.Subject = "חוויות מטיול פגסוס"
                Msg.To = Email
                Msg.Body = Email_content.ToString()
                System.Web.Mail.SmtpMail.Send(Msg)
                Msg = Nothing
            End If
            Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO Email_to_contact (" & _
           " company_id,contact_id,Traveler_Id,Email_Type_id,contact_Email,FromEmail,Email_Content,User_id,departure_id,date_send)" & _
           " VALUES (" & companyID & "," & ContactId & "," & TravelerId & "," & Emailtype_id & ",'" & Email & "','" & FromEmail & "','" & Email_content & "'," & UserID & "," & DepartureId & ",GetDate())", conB)
            cmdInsert.CommandType = CommandType.Text
            '   cmdInsert.Parameters.Add("@region_name", SqlDbType.VarChar, 150).Value = CStr(regionName)
            '    cmdInsert.Parameters.Add("@max_order", SqlDbType.Int).Value = MaxOrder
            conB.Open()
            cmdInsert.ExecuteNonQuery()
            conB.Close()

            Dim cmdInsertT As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO Email_to_Traveler (" & _
 " contact_id,Traveler_Id,Email_Type_id,contact_Email,FromEmail,Email_Content,User_id,departure_id,date_send)" & _
 " VALUES (" & ContactId & "," & TravelerId & "," & Emailtype_id & ",'" & Email & "','" & FromEmail & "','" & Email_content & "'," & UserID & "," & DepartureId & ",GetDate())", conB)
            cmdInsertT.CommandType = CommandType.Text
            '   cmdInsert.Parameters.Add("@region_name", SqlDbType.VarChar, 150).Value = CStr(regionName)
            '    cmdInsert.Parameters.Add("@max_order", SqlDbType.Int).Value = MaxOrder
            conB.Open()
            cmdInsertT.ExecuteNonQuery()
            conB.Close()





            ' Response.Write(Email & "-" & Email_content & "<BR>")
            '   Response.End()
            'send ....

        End While
        dr_m.Close()
        cmdSelect.Dispose()
        conPegasus.Close()

        Response.Write("send mail- end ")

    End Sub

End Class

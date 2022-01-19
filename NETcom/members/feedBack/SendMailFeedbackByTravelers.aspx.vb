Imports System.Data.SqlClient
Public Class SendMailFeedbackByTravelers
    Inherits System.Web.UI.Page
    Protected DepartureId As Integer = 0
    Protected conBU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected drpEmail_type_id As Web.UI.WebControls.DropDownList
    Protected EmailText As Double
    Protected Email_content As String
    Protected Departure_Code, Date_Begin, Date_End, EmailSubject, ContactName As String
    Protected func As New include.funcs
    Dim cmdSelect, cmdSelect1 As SqlClient.SqlCommand
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionStringS"))
    Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionStringS"))
    Dim dr, drOrg As SqlClient.SqlDataReader
    Public dr_m As SqlClient.SqlDataReader
    Public TravelersId As String
    Public CountEmail As Array

    Public SiteId, ContactId, TravelerId, EmailStatusId, companyID, UserID As Integer
    Public siteName, addLang, dirText, Email, FromEmail, Emailtype_id, orgName, DateBegin, TourTitle, GuideName, FeedBack_Url, FeedBack_content As String


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
        'CountEmail = 1
        orgName = "pegasus"
        If IsNumeric(Request.QueryString("companyId")) Then
            companyID = CInt(Request.QueryString("companyId"))
        Else
            companyID = 0
        End If
        If IsNumeric(Request.QueryString("sCheck")) Then
            TravelersId = Request.QueryString("sCheck")
        Else
            TravelersId = "0"
        End If

        'SiteId = func.getSite(Page)
        'If SiteId = 1 Then
        '    func.Chk_Permiton(3)
        'ElseIf SiteId = 2 Then
        '    func.Chk_Permiton(14)
        'End If
        If IsNumeric(Request.QueryString("sdep")) Then
            DepartureId = CInt(Request.QueryString("sdep"))
        Else
            DepartureId = 0
        End If
        UserID = 1028   'מנהל אתר אינטרנט



        ' Response.Write("DepartureId=" & DepartureId & "<BR>")
        ' Response.Write("TravelersId=" & TravelersId & "<BR>")
        ' Response.End()
        'siteName = func.dt_sites.Rows.Find(SiteId)("site_Name")
        'addLang = func.dt_sites.Rows.Find(SiteId)("Field_Addition")
        'dirText = func.dt_sites.Rows.Find(SiteId)("Text_Direction")

        Emailtype_id = 8   'משוב טיול
        Emailtype_id = 88   'משוב מטייל 
        CountEmail = TravelersId.Split(",")
        'Response.Write("g=" & TravelersId)
        'Response.Write("count=" & CountEmail.Length)
        'Response.End()

        If Request.Form.Count = 0 Then

            Dim cmdTmp = New SqlClient.SqlCommand("select Date_Begin,Tour_Title, Guide_FName +' ' + Guide_LName  as Guide_Name  from Tours_Departures " & _
             " left join tours on tours.tour_id=Tours_Departures.tour_id left join  Guides on Guides.Guide_Id=Tours_Departures.Guide_Id where(Departure_Id = @DepartureId)", con)
            cmdTmp.CommandType = CommandType.Text
            cmdTmp.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
            con.Open()
            drOrg = cmdTmp.ExecuteReader(CommandBehavior.SingleRow)
            If drOrg.Read() Then
                If Not drOrg("Date_Begin") Is DBNull.Value Then
                    DateBegin = Trim(drOrg("Date_Begin"))
                Else
                    DateBegin = ""
                End If
                If Not drOrg("Tour_Title") Is DBNull.Value Then
                    TourTitle = Trim(drOrg("Tour_Title"))
                Else
                    TourTitle = ""
                End If
                If Not drOrg("Guide_Name") Is DBNull.Value Then
                    GuideName = Trim(drOrg("Guide_Name"))
                Else
                    GuideName = ""
                End If
            End If


            FromEmail = "info@pegasusisrael.co.il"


            con.Close()



        Else

            FromEmail = Request.Form("FromEmail")
            Emailtype_id = 8 'שליחת משוב drpEmail_type_id.SelectedValue
            Emailtype_id = 88 'שליחת משוב למטייל
            '   Response.Write("t=" & Emailtype_id)
            '  Response.End()

            ''''''--- Email_content = Trim(Request.Form("Email_content"))

            '   Dim MemberEmail, FullName, login, password As String
            If TravelersId <> "0" Then
                'Dim s = "SELECT  TT.Contact_Id,TT.Traveler_Id,FirstName+''+LastName as CONTACT_NAME,TT.Email,12681 as Company_Id FROM Tour_Travelers TT " & _
                '    " left join Travelers T on TT.Traveler_Id=T.Traveler_Id  WHERE (Departure_Id =" & DepartureId & ") " & _
                '    " and TT.Traveler_Id  in (" & TravelersId & " ) and len(Email)>0"
                'Response.Write(s)
                'Response.End()
                cmdSelect = New SqlCommand("SELECT  TT.Contact_Id,TT.Traveler_Id,FirstName+' '+LastName as CONTACT_NAME,TT.Email,12681 as Company_Id FROM Tour_Travelers TT " & _
                    " left join Travelers T on TT.Traveler_Id=T.Traveler_Id  WHERE (Departure_Id =" & DepartureId & ") " & _
                    " and TT.Traveler_Id  in (" & TravelersId & " ) and len(Email)>0", conBU)
                ' cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
                'cmdSelect.Parameters.Add("@TravelersId", SqlDbType.NVarChar, 500).Value = CInt(TravelersId)
                cmdSelect.CommandType = CommandType.Text
                conBU.Open()

            End If
            dr_m = cmdSelect.ExecuteReader()

            While dr_m.Read()
                If Not dr_m("Email") Is DBNull.Value Then
                    Email = Trim(dr_m("Email"))
                    ' Email = "furfaina@gmail.com"
                End If
                If Not dr_m("Contact_Id") Is DBNull.Value Then
                    ContactId = Trim(dr_m("Contact_Id"))
                End If
                If Not dr_m("Traveler_Id") Is DBNull.Value Then
                    TravelerId = Trim(dr_m("Traveler_Id"))
                End If
                If Not dr_m("CONTACT_NAME") Is DBNull.Value Then
                    ContactName = dr_m("CONTACT_NAME")

                Else
                    ContactName = ""

                End If

                '  Response.End()



                If Not dr_m("Company_Id") Is DBNull.Value Then
                    companyID = Trim(dr_m("Company_Id"))
                End If

                If Trim(Email) <> "" Then


                    FeedBack_Url = ConfigurationSettings.AppSettings.Item("PegasusFeedbackUrl") & "/feedback/faqIsTraveler.aspx?DepartureId=" & DepartureId & "&TravelerId=" & TravelerId & "&ContactId=" & ContactId
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
                    FeedBack_content = ContactName & ","
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
                    '   ---block send Email---

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



                End If


            End While
            dr_m.Close()
            cmdSelect.Dispose()
            conBU.Close()
            Response.Write("<script language='javascript'> {window.opener.location.href = window.opener.location.href; window.close();}</script>")

        End If

    End Sub

End Class
Imports System.Data.SqlClient
Public Class SendMailFeedBackIsTraveler
    Inherits System.Web.UI.Page
    Protected DepartureId As Integer = 0
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected conBU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected drpEmail_type_id As Web.UI.WebControls.DropDownList
    Protected EmailText As Double
    Protected Email_content As String
    Protected Departure_Code, Date_Begin, Date_End, EmailSubject, TravelerName, FromEmail As String
    Protected func As New include.funcs
    Dim cmdSelect, cmdSelect1 As SqlClient.SqlCommand
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionStringS"))
    Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionStringS"))
    Dim dr, drOrg As SqlClient.SqlDataReader
    Public dr_m As SqlClient.SqlDataReader
    Public SiteId, CountMail, ContactId, TravelerId, EmailStatusId, companyID, UserID As Integer
    Public siteName, addLang, dirText, Email, Email_phone, Emailtype_id, orgName, DateBegin, TourTitle, GuideName, FeedBack_Url, FeedBack_content As String


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
        CountMail = 1
        orgName = "pegasus"
        If IsNumeric(Request.QueryString("companyId")) Then
            companyID = CInt(Request.QueryString("companyId"))
        Else
            companyID = 0
        End If
        companyID = 12681
        If IsNumeric(Request.QueryString("ContactId")) Then
            ContactId = CInt(Request.QueryString("ContactId"))
        Else
            ContactId = 0
        End If

        If IsNumeric(Request.QueryString("TravelerId")) Then
            TravelerId = CInt(Request.QueryString("TravelerId"))
        Else
            TravelerId = 0
        End If
        If IsNumeric(Request.QueryString("DepartureId")) Then
            DepartureId = CInt(Request.QueryString("DepartureId"))
        Else
            DepartureId = 0
        End If
        UserID = 1028   '���� ��� �������

        FromEmail = "info@pegasusisrael.co.il"


        Emailtype_id = 8   '���� ����
        Emailtype_id = 88   ' ���� ���� ������
        '' Response.Write("1")
        If Request.Form.Count = 0 Then

            '    Response.Write("Request.Form.Count")
            '      Response.End()
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



            con.Close()

        Else
            CountMail = 1
            '  Response.Write("count=1")
            'Email_phone = Request.Form("Email_phone")
            Emailtype_id = 8 '����� ���� drpEmail_type_id.SelectedValue
            Emailtype_id = 88   ' ���� ���� ������
            If TravelerId <> 0 Then
                cmdSelect = New SqlCommand("GetContactMailFeedBackByTraveler", conBU)
                cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
                cmdSelect.Parameters.Add("@TravelerId", SqlDbType.Int).Value = CInt(TravelerId)
                cmdSelect.CommandType = CommandType.StoredProcedure
                conBU.Open()

            Else
                cmdSelect = New SqlCommand("GetContactMailFeedBackIsTraveler", conBU)
                cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
                cmdSelect.CommandType = CommandType.StoredProcedure
                conBU.Open()

            End If
            dr_m = cmdSelect.ExecuteReader()
            '  Response.Write("cmdSelect=" & cmdSelect.CommandText)
            While dr_m.Read()
                If Not dr_m("Email") Is DBNull.Value Then
                    'Email = Trim(dr_m("Email"))
                    Email = "faina@cyberserve.co.il"
                End If
                If Not dr_m("Contact_Id") Is DBNull.Value Then
                    ContactId = Trim(dr_m("Contact_Id"))
                End If
                If Not dr_m("Traveler_Id") Is DBNull.Value Then
                    TravelerId = dr_m("Traveler_Id")
                End If
                'If Not dr_m("CONTACT_NAME") Is DBNull.Value Then
                '    ContactName = dr_m("CONTACT_NAME")

                'Else
                '    ContactName = ""

                'End If
                If Not dr_m("TravelerName") Is DBNull.Value Then
                    TravelerName = dr_m("TravelerName")

                Else
                    TravelerName = ""

                End If

                '  Response.End()

                'If Not dr_m("Member_Id") Is DBNull.Value Then
                '    MemberId = dr_m("Member_Id")
                'End If

                'If Not dr_m("Company_Id") Is DBNull.Value Then
                '    companyID = Trim(dr_m("Company_Id"))
                'End If


                'Response.Write(DepartureId & "-" & ContactId & ":" & Email & "<BR>")
                'Response.End()

                If Trim(Email) <> "" Then
                    '      Response.Write("Email=" & Email)
                    FeedBack_Url = ConfigurationSettings.AppSettings.Item("PegasusFeedbackUrl") & "/feedback/FaqIsTraveler.aspx?Email=1&DepartureId=" & DepartureId & "&ContactId=" & ContactId & "&TravelerId=" & TravelerId
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

                    ' Response.Write("urlF=" & urlF)
                    ' Response.End()
                    xmlhttpShort = Nothing
                    FeedBack_content = TravelerName & "<BR>"
                    FeedBack_content = FeedBack_content & "<BR> ���� ���� ���� <BR>���� ����� ���� �� ����� ������ �� ������ <BR>"
                    FeedBack_content = FeedBack_content & urlF
                    FeedBack_content = FeedBack_content & "<BR>" & " ����, ����� "
                    Email_content = FeedBack_content

                    '  If Trim(Email) <> "" Then
                    Dim Msg As New System.Web.Mail.MailMessage
                    Msg.BodyFormat = System.Web.Mail.MailFormat.Html
                    Msg.BodyEncoding = System.Text.Encoding.UTF8
                    Msg.From = "info@pegasusisrael.co.il"
                    Msg.Subject = "������ ����� �����"
                    Msg.To = Email
                    Msg.Body = Email_content.ToString()
                    System.Web.Mail.SmtpMail.Send(Msg)
                    Msg = Nothing
                    ' End If


                    'Response.Write("send")
                    '   Response.End()

                    Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO Email_to_Traveler (" & _
    " contact_id,Traveler_Id,Email_Type_id,contact_Email,FromEmail,Email_Content,User_id,departure_id,date_send)" & _
    " VALUES (" & ContactId & "," & TravelerId & "," & Emailtype_id & ",'" & Email & "','" & FromEmail & "','" & Email_content & "'," & UserID & "," & DepartureId & ",GetDate())", conB)
                    cmdInsert.CommandType = CommandType.Text
                    '   cmdInsert.Parameters.Add("@region_name", SqlDbType.VarChar, 150).Value = CStr(regionName)
                    '    cmdInsert.Parameters.Add("@max_order", SqlDbType.Int).Value = MaxOrder
                    conB.Open()
                    cmdInsert.ExecuteNonQuery()
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
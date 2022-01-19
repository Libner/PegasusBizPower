Imports System.Data.SqlClient
Public Class SendSMSFeedBackIsTraveler
    Inherits System.Web.UI.Page
    Protected DepartureId As Integer = 0
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected drpSMS_type_id As Web.UI.WebControls.DropDownList
    Protected smsText As Double
    Protected SMS_content As String
    Protected Departure_Code, Date_Begin, Date_End, EmailSubject, TravelerName As String
    Protected func As New include.funcs
    Dim cmdSelect, cmdSelect1 As SqlClient.SqlCommand
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionStringS"))
    Dim conBU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionString"))
    Dim dr, drOrg As SqlClient.SqlDataReader
    Public dr_m As SqlClient.SqlDataReader
    Public SiteId, CountSms, ContactId, TravelerId, MemberId, smsStatusId, companyID, UserID As Integer
    Public Phone, siteName, addLang, dirText, sms_phone, SMStype_id, orgName, DateBegin, TourTitle, GuideName, FeedBack_Url, FeedBack_content As String


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
        CountSms = 1
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

        'SiteId = func.getSite(Page)
        'If SiteId = 1 Then
        '    func.Chk_Permiton(3)
        'ElseIf SiteId = 2 Then
        '    func.Chk_Permiton(14)
        'End If
        If IsNumeric(Request.QueryString("DepartureId")) Then
            DepartureId = CInt(Request.QueryString("DepartureId"))
        Else
            DepartureId = 0
        End If
        UserID = 1028   'מנהל אתר אינטרנט



        SMStype_id = 8   'משוב טיול
        SMStype_id = 88   ' משוב טיול למטייל

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


            sms_phone = "036374000"


            con.Close()
            '  If ContactId = 0 Then
            '    Dim cmdTmpCount = New SqlClient.SqlCommand("SELECT Count(*) as CountSMS FROM dbo.Members WHERE (Departure_Id = @DepartureId)", con)
            '    cmdTmpCount.CommandType = CommandType.Text
            '    cmdTmpCount.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
            '    con.Open()
            '    dr = cmdTmpCount.ExecuteReader(CommandBehavior.SingleRow)
            '    If dr.Read() Then
            '        CountSms = dr("CountSMS")
            '    End If
            '    con.Close()
            'End If


        Else
            CountSms = 1
          
            sms_phone = Request.Form("sms_phone")
            SMStype_id = 8 'שליחת משוב drpSMS_type_id.SelectedValue
            SMStype_id = 88   ' משוב טיול למטייל
            '   Response.Write("t=" & SMStype_id)
            '  Response.End()

            ''''''--- SMS_content = Trim(Request.Form("SMS_content"))

            '   Dim MemberEmail, FullName, login, password As String
            If TravelerId <> 0 Then
                cmdSelect = New SqlCommand("GetContactSMSFeedBackByTraveler", conBU)
                cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
                cmdSelect.Parameters.Add("@TravelerId", SqlDbType.Int).Value = CInt(TravelerId)
                cmdSelect.CommandType = CommandType.StoredProcedure
                conBU.Open()

            Else
                cmdSelect = New SqlCommand("GetContactSMSFeedBackIsTraveler", conBU)
                cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
                cmdSelect.CommandType = CommandType.StoredProcedure
                conBU.Open()

            End If
            'cmdSelect = New SqlCommand("GetContactSMSFeedBack", con)
            'cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
            'cmdSelect.CommandType = CommandType.StoredProcedure
            'con.Open()
            dr_m = cmdSelect.ExecuteReader()

            While dr_m.Read()
                If Not dr_m("Phone") Is DBNull.Value Then
                    Phone = dr_m("Phone")
                    'Phone = "0507740302"
                End If
                If Not dr_m("Traveler_Id") Is DBNull.Value Then
                    TravelerId = dr_m("Traveler_Id")
                End If

                If Not dr_m("Contact_Id") Is DBNull.Value Then
                    ContactId = Trim(dr_m("Contact_Id"))
                End If
                If Not dr_m("TravelerName") Is DBNull.Value Then
                    TravelerName = dr_m("TravelerName")

                Else
                    TravelerName = ""

                End If
                '  Response.Write("TravelerId=" & TravelerId)
                '  Response.End()

                'If Not dr_m("Member_Id") Is DBNull.Value Then
                '    MemberId = dr_m("Member_Id")
                'End If

                'If Not dr_m("Company_Id") Is DBNull.Value Then
                '    companyID = Trim(dr_m("Company_Id"))
                'End If


             
                If Trim(Phone) <> "" Then
                    If Len(Phone) < 10 Or Len(Phone) > 10 Then
                        smsStatusId = 3 ' מספר שגוי
                        Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_Travelers (" & _
                                " Traveler_Id,sms_Type_id,sms_Status_Id,contact_Id,Traveler_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
                                " VALUES (" & TravelerId & "," & SMStype_id & "," & ContactId & "," & smsStatusId & ",'" & Phone & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate()," & DepartureId & ")", conB)

                        cmdInsert.CommandType = CommandType.Text
                        conB.Open()
                        cmdInsert.ExecuteNonQuery()
                        conB.Close()
                        Dim cmdInsert2 As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_Contact (" & _
                      " Traveler_Id,sms_Type_id,sms_Status_Id,contact_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
                      " VALUES (" & TravelerId & "," & SMStype_id & "," & smsStatusId & "," & ContactId & ",'" & Phone & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate()," & DepartureId & ")", conB)

                        cmdInsert2.CommandType = CommandType.Text
                        conB.Open()
                        cmdInsert2.ExecuteNonQuery()
                        conB.Close()


                    ElseIf Left(Phone, 3) <> "050" And Left(Phone, 3) <> "052" And Left(Phone, 3) <> "053" And Left(Phone, 3) <> "054" And Left(Phone, 3) <> "055" And Left(Phone, 3) <> "058" And Left(Phone, 3) <> "077" Then
                        smsStatusId = 3
                        Dim str
                        Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_Travelers (" & _
            " Traveler_Id,sms_Type_id,sms_Status_Id,contact_Id,Traveler_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
            " VALUES (" & TravelerId & "," & SMStype_id & "," & smsStatusId & "," & ContactId & ",'" & Phone & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate()," & DepartureId & ")", conB)
                        cmdInsert.CommandType = CommandType.Text
                        conB.Open()
                        cmdInsert.ExecuteNonQuery()
                        conB.Close()
                        Dim cmdInsert2 As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_Contact (" & _
                  " Traveler_Id,sms_Type_id,sms_Status_Id,contact_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
                  " VALUES (" & TravelerId & "," & SMStype_id & "," & smsStatusId & "," & ContactId & ",'" & Phone & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate()," & DepartureId & ")", conB)
                        cmdInsert2.CommandType = CommandType.Text
                        conB.Open()
                        cmdInsert2.ExecuteNonQuery()
                        conB.Close()

                    Else


                        FeedBack_Url = ConfigurationSettings.AppSettings.Item("PegasusFeedbackUrl") & "/feedback/faqIsTraveler.aspx?DepartureId=" & DepartureId & "&TravelerId=" & TravelerId & "&ContactId=" & ContactId
                        FeedBack_Url = "url=" & Server.UrlEncode(FeedBack_Url)
                        'Response.Write("FeedBack_Url=" & FeedBack_Url)
                        'Response.End()

                        Dim xmlhttpShort As Object
                        Dim sendUrlShort, strResponseShort, urlF As String
                        sendUrlShort = "http://tinyurl.com/api-create.php"
                        xmlhttpShort = CreateObject("MSXML2.ServerXMLHTTP")
                        xmlhttpShort.open("POST", sendUrlShort, False)

                        xmlhttpShort.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
                        xmlhttpShort.send(FeedBack_Url & "&charset=iso-8859-8")
                        strResponseShort = xmlhttpShort.responseText
                        urlF = strResponseShort

                        xmlhttpShort = Nothing

                        FeedBack_content = TravelerName & ","
                        FeedBack_content = FeedBack_content & Chr(13)
                        FeedBack_content = FeedBack_content & " ברוך שובך ארצה. נשמח לקבלת משוב על הטיול בלחיצה על הקישור :"
                        FeedBack_content = FeedBack_content & urlF
                        FeedBack_content = FeedBack_content & " תודה, פגסוס "
                        SMS_content = FeedBack_content
                        '  Response.Write(FeedBack_content)
                        ' Response.End()

                        smsStatusId = 0
                        Dim sendUrl, strResponse, getUrl As String
                        sendUrl = "http://www.micropay.co.il/ExtApi/ScheduleSms.php"

                        Dim xmlhttp As Object
                        xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
                        xmlhttp.open("POST", sendUrl, False)
                        ' ---block send sms---
                        xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
                        xmlhttp.send("uid=2575&un=pegasus&msglong=" & Server.UrlEncode(SMS_content) & "&charset=iso-8859-8" & _
                       "&from=" & sms_phone & "&post=2&list=" & Server.UrlEncode(Phone) & _
                       "&desc=" & Server.UrlEncode(orgName))
                        strResponse = xmlhttp.responseText
                        xmlhttp = Nothing

                        '   ---block send sms---

                        If InStr(strResponse, "ERROR") > 0 Then
                            smsStatusId = 2
                        ElseIf InStr(strResponse, "OK") > 0 Then
                            smsStatusId = 1

                        End If
                        '  Response.Write("<BR>" & "smsStatusId=" & smsStatusId & "<BR>" & Server.UrlEncode(Phone))
                        '  Response.End()
                        Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_Travelers (" & _
    " Traveler_Id,sms_Type_id,sms_Status_Id,contact_Id,Traveler_CELL,sms_phone,sms_Content,User_id,departure_id,date_send)" & _
    " VALUES (" & TravelerId & "," & SMStype_id & "," & smsStatusId & "," & ContactId & ",'" & Phone & "','" & sms_phone & "','" & SMS_content & "'," & UserID & "," & DepartureId & ",GetDate())", conB)
                        cmdInsert.CommandType = CommandType.Text
                        conB.Open()
                        cmdInsert.ExecuteNonQuery()
                        conB.Close()
                        Dim cmdInsert2 As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_Contact (" & _
                          " Traveler_Id,sms_Type_id,sms_Status_Id,contact_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
                          " VALUES (" & TravelerId & "," & SMStype_id & "," & smsStatusId & "," & ContactId & ",'" & Phone & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate()," & DepartureId & ")", conB)
                        cmdInsert2.CommandType = CommandType.Text
                        conB.Open()
                        cmdInsert2.ExecuteNonQuery()
                        conB.Close()
                        Dim cmdSelectUp As New SqlClient.SqlCommand("SendSMSFeedBack_IsTravelers", conB)
                        cmdSelectUp.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
                        cmdSelectUp.Parameters.Add("@TravelerId", SqlDbType.Int).Value = CInt(TravelerId)
                        cmdSelectUp.Parameters.Add("@ContactId", SqlDbType.Int).Value = CInt(ContactId)
                        cmdSelectUp.Parameters.Add("@smsStatusId", SqlDbType.Int).Value = CInt(smsStatusId)
                        cmdSelectUp.Parameters.Add("@Cellular", SqlDbType.VarChar, 50).Value = Phone
                        cmdSelectUp.Parameters.Add("@sms_phone", SqlDbType.VarChar, 50).Value = sms_phone
                        cmdSelectUp.Parameters.Add("@FeedBack_Url", SqlDbType.VarChar, 500).Value = urlF 'FeedBack_Url
                        cmdSelectUp.Parameters.Add("@FeedBack_content", SqlDbType.VarChar, 500).Value = FeedBack_content
                        cmdSelectUp.Parameters.Add("@UserID", SqlDbType.Int).Value = CInt(UserID)
                        cmdSelectUp.CommandType = CommandType.StoredProcedure
                        conB.Open()
                        cmdSelectUp.ExecuteNonQuery()
                        conB.Close()

                    End If
                End If
            End While
            dr_m.Close()
            cmdSelect.Dispose()
            conBU.Close()
            Response.Write("<script language='javascript'> {window.opener.location.href = window.opener.location.href; window.close();}</script>")

        End If

    End Sub

End Class
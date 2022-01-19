Imports System.Data.SqlClient
Public Class SendSMSFeedbackByContacts
    Inherits System.Web.UI.Page
    Protected DepartureId As Integer = 0
    Protected conB As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings.Item("ConnectionString"))
    Protected drpSMS_type_id As Web.UI.WebControls.DropDownList
    Protected smsText As Double
    Protected SMS_content As String
    Protected Departure_Code, Date_Begin, Date_End, EmailSubject, ContactName As String
    Protected func As New include.funcs
    Dim cmdSelect, cmdSelect1 As SqlClient.SqlCommand
    Dim con As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionStringS"))
    Dim conU As New SqlClient.SqlConnection(ConfigurationSettings.AppSettings("ConnectionStringS"))
    Dim dr, drOrg As SqlClient.SqlDataReader
    Public dr_m As SqlClient.SqlDataReader
    Public ContactsId As String
    Public CountSms As Array

    Public SiteId, ContactId, MemberId, smsStatusId, companyID, UserID As Integer
    Public siteName, addLang, dirText, Cellular, sms_phone, SMStype_id, orgName, DateBegin, TourTitle, GuideName, FeedBack_Url, FeedBack_content As String


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
        'CountSms = 1
        orgName = "pegasus"
        If IsNumeric(Request.QueryString("companyId")) Then
            companyID = CInt(Request.QueryString("companyId"))
        Else
            companyID = 0
        End If
        If IsNumeric(Request.QueryString("sCheck")) Then
            ContactsId = Request.QueryString("sCheck")
        Else
            ContactsId = "0"
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
        ' Response.Write("ContactsId=" & ContactsId & "<BR>")
        ' Response.End()
        'siteName = func.dt_sites.Rows.Find(SiteId)("site_Name")
        'addLang = func.dt_sites.Rows.Find(SiteId)("Field_Addition")
        'dirText = func.dt_sites.Rows.Find(SiteId)("Text_Direction")

        SMStype_id = 8   'משוב טיול
        CountSms = ContactsId.Split(",")
        'Response.Write("g=" & ContactsId)
        'Response.Write("count=" & CountSms.Length)
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


            sms_phone = "036374000"


            con.Close()
         


        Else
          
            sms_phone = Request.Form("sms_phone")
            SMStype_id = 8 'שליחת משוב drpSMS_type_id.SelectedValue
            '   Response.Write("t=" & SMStype_id)
            '  Response.End()

            ''''''--- SMS_content = Trim(Request.Form("SMS_content"))

            '   Dim MemberEmail, FullName, login, password As String
            If ContactsId <> "0" Then
                ' cmdSelect = New SqlCommand("GetContactSMSFeedBackByContacts", con)


                cmdSelect = New SqlCommand("SELECT  M.Member_Id, M.Contact_Id,C.CONTACT_NAME,M.Cellular,Company_Id FROM Members M left join bizpower_pegasus.dbo.CONTACTS C on M.Contact_id=C.Contact_Id  " & _
             " WHERE (Departure_Id = " & DepartureId & ") and M.Contact_Id  in (" & ContactsId & ")", con)
                ' cmdSelect.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
                'cmdSelect.Parameters.Add("@ContactsId", SqlDbType.NVarChar, 500).Value = CInt(ContactsId)
                cmdSelect.CommandType = CommandType.Text
                con.Open()

            End If
            dr_m = cmdSelect.ExecuteReader()

            While dr_m.Read()
                If Not dr_m("Cellular") Is DBNull.Value Then
                    Cellular = Trim(dr_m("Cellular"))
                    ' Cellular = "0507740302"
                End If
                If Not dr_m("Contact_Id") Is DBNull.Value Then
                    ContactId = Trim(dr_m("Contact_Id"))
                End If
                If Not dr_m("CONTACT_NAME") Is DBNull.Value Then
                    ContactName = dr_m("CONTACT_NAME")

                Else
                    ContactName = ""

                End If

                '  Response.End()

                If Not dr_m("Member_Id") Is DBNull.Value Then
                    MemberId = dr_m("Member_Id")
                End If

                If Not dr_m("Company_Id") Is DBNull.Value Then
                    companyID = Trim(dr_m("Company_Id"))
                End If


                '  Response.Write(DepartureId & "-" & ContactId & ":" & Cellular & "<BR>")
                ' Response.End()

                If Trim(Cellular) <> "" Then
                    If Len(Cellular) < 10 Then
                        smsStatusId = 3 ' מספר שגוי
                        Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_contact (" & _
                 " company_id,contact_id,sms_Type_id,sms_Status_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
                 " VALUES (" & companyID & "," & ContactId & "," & SMStype_id & "," & smsStatusId & ",'" & Cellular & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate())," & DepartureId, conB)
                        cmdInsert.CommandType = CommandType.Text
                        '   cmdInsert.Parameters.Add("@region_name", SqlDbType.VarChar, 150).Value = CStr(regionName)
                        '    cmdInsert.Parameters.Add("@max_order", SqlDbType.Int).Value = MaxOrder
                        conB.Open()
                        cmdInsert.ExecuteNonQuery()
                        conB.Close()

                    ElseIf Left(Cellular, 3) <> "050" And Left(Cellular, 3) <> "052" And Left(Cellular, 3) <> "053" And Left(Cellular, 3) <> "054" And Left(Cellular, 3) <> "055" And Left(Cellular, 3) <> "058" And Left(Cellular, 3) <> "077" Then
                        smsStatusId = 3
                        Dim str
                        '        str = "SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_contact (" & _
                        '" company_id,contact_id,sms_Type_id,sms_Status_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
                        '" VALUES (" & companyID & "," & ContactId & "," & SMStype_id & "," & smsStatusId & ",'" & Cellular & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate()," & DepartureId & ")"
                        '      Response.Write(str)
                        '     Response.End()
                        Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_contact (" & _
                 " company_id,contact_id,sms_Type_id,sms_Status_Id,contact_CELL,sms_phone,sms_Content,User_id,date_send,Departure_Id)" & _
                 " VALUES (" & companyID & "," & ContactId & "," & SMStype_id & "," & smsStatusId & ",'" & Cellular & "','" & sms_phone & "','" & SMS_content & "'," & UserID & ",GetDate()," & DepartureId & ")", conB)
                        cmdInsert.CommandType = CommandType.Text
                        '   cmdInsert.Parameters.Add("@region_name", SqlDbType.VarChar, 150).Value = CStr(regionName)
                        '    cmdInsert.Parameters.Add("@max_order", SqlDbType.Int).Value = MaxOrder
                        conB.Open()
                        cmdInsert.ExecuteNonQuery()
                        conB.Close()

                    Else

                        'SMS_content = ContactName
                        'SMS_content = SMS_content & "http://www.pegasusisrael.co.il/feedback/faq.aspx?ContactId"
                        '  FeedBack_Url = "http://194.90.203.114/feedback/faq.aspx?ContactId=" & ContactId & "&departureid=" & DepartureId
                        ' FeedBack_content = "קיים משוב"
                        'SMS_content = SMS_content & "http://194.90.203.114/feedback/faq.aspx?ContactId=" & ContactId & "&departureid=" & DepartureId
                        ' SMS_content = SMS_content & " שלום רב,אנו מקווים כי נהנית ומודים לך על השתתפותך בטיול " & TourTitle & " בהדרכת " & GuideName & " שיצא בתאריך " & DateBegin & ",. מתוך רצון כנה להקשיב לך ובמטרה לקדם ולשפר את המוצר והשירות שניתן לך נשמח לשיתוף פעולה במילוי משוב קצר. חברת פגסוס רואה במשוב זה אמצעי חשוב ביותר לקידום ושיפור השירות ומתייחסת במלוא הרצינות וכובד ראש לממצאים העולים"
                        ' SMS_content = SMS_content & "אנו מודים על נכונותך להקדיש מעט מזמנך ולענות למשוב זה: htttp://www.pegasusisrael.co.il/feedback/faq.aspx?ContactId=" & ContactId & "&departureid=" & DepartureId

                        FeedBack_Url = "https://www.pegasusisrael.co.il/feedback/faq.aspx?DepartureId=" & DepartureId & "&ContactId=" & ContactId
                        FeedBack_Url = "url=" & Server.UrlEncode(FeedBack_Url)
                        'FeedBack_Url = "url=http://www.pegasusisrael.co.il/feedback/faq.aspx?DepartureId=" & DepartureId & "&ContactId=" & ContactId

                        Response.Write("FeedBack_Url=" & FeedBack_Url & "<BR>")
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
                        'If Err.Number = 0 Then
                        '    urlF = urlF
                        'Else
                        '    urlF = FeedBack_Url
                        '    'End If
                        'Response.Write("xmlhttpShort=" & xmlhttpShort.responseText & "<BR>")
                        'Response.Write("Err.Number =" & Err.Number & "<BR>")
                        'Response.Write("Err.Description=" & Err.Description & "<BR>")

                        'Response.End()
                        xmlhttpShort = Nothing
                        FeedBack_content = " ברוך שובך ארצה. נשמח לקבלת משוב על הטיול בלחיצה על הקישור :"
                        FeedBack_content = FeedBack_content & urlF
                        FeedBack_content = FeedBack_content & " תודה, פגסוס "
                        SMS_content = FeedBack_content

                        smsStatusId = 0
                        Dim sendUrl, strResponse, getUrl As String
                        sendUrl = "http://www.micropay.co.il/ExtApi/ScheduleSms.php"

                        Dim xmlhttp As Object
                        xmlhttp = CreateObject("MSXML2.ServerXMLHTTP")
                        xmlhttp.open("POST", sendUrl, False)
                        ' ---block send sms---
                        xmlhttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded")
                        xmlhttp.send("uid=2575&un=pegasus&msglong=" & Server.UrlEncode(SMS_content) & "&charset=iso-8859-8" & _
                     "&from=" & sms_phone & "&post=2&list=" & Server.UrlEncode(Cellular) & _
                       "&desc=" & Server.UrlEncode(orgName))
                        strResponse = xmlhttp.responseText
                        xmlhttp = Nothing
                        '   ---block send sms---

                        If InStr(strResponse, "ERROR") > 0 Then
                            'error = Mid(strResponse, 6)	
                            smsStatusId = 2
                        ElseIf InStr(strResponse, "OK") > 0 Then
                            smsStatusId = 1
                            '   tid = Trim(Mid(strResponse, 3))
                            '  If IsNumeric(tid) Then
                            ' tid = CLng(tid)
                            ' Else
                            '     tid = 0
                            ' End If
                        End If
                        Dim cmdInsert As New SqlClient.SqlCommand("SET DATEFORMAT DMY; SET NOCOUNT ON; INSERT INTO sms_to_contact (" & _
    " company_id,contact_id,sms_Type_id,sms_Status_Id,contact_CELL,sms_phone,sms_Content,User_id,departure_id,date_send)" & _
    " VALUES (" & companyID & "," & ContactId & "," & SMStype_id & "," & smsStatusId & ",'" & Cellular & "','" & sms_phone & "','" & SMS_content & "'," & UserID & "," & DepartureId & ",GetDate())", conB)
                        cmdInsert.CommandType = CommandType.Text
                        '   cmdInsert.Parameters.Add("@region_name", SqlDbType.VarChar, 150).Value = CStr(regionName)
                        '    cmdInsert.Parameters.Add("@max_order", SqlDbType.Int).Value = MaxOrder
                        conB.Open()
                        cmdInsert.ExecuteNonQuery()
                        conB.Close()
                        ' Response.Write("Cellular=" & Cellular & "<BR>")
                        ' Response.Write("FeedBack_Url=" & FeedBack_Url)
                        ' Response.Write("urlF=" & urlF & "<BR>")
                        ' Response.End()
                        Dim cmdSelectUp = New SqlCommand("SendSMSFeedBack", conB)
                        cmdSelectUp.Parameters.Add("@DepartureId", SqlDbType.Int).Value = CInt(DepartureId)
                        cmdSelectUp.Parameters.Add("@companyID", SqlDbType.Int).Value = CInt(companyID)
                        cmdSelectUp.Parameters.Add("@ContactId", SqlDbType.Int).Value = CInt(ContactId)
                        cmdSelectUp.Parameters.Add("@smsStatusId", SqlDbType.Int).Value = CInt(smsStatusId)
                        cmdSelectUp.Parameters.Add("@Cellular", SqlDbType.VarChar, 50).Value = Cellular
                        cmdSelectUp.Parameters.Add("@sms_phone", SqlDbType.VarChar, 50).Value = sms_phone
                        cmdSelectUp.Parameters.Add("@FeedBack_Url", SqlDbType.VarChar, 500).Value = urlF 'FeedBack_Url
                        cmdSelectUp.Parameters.Add("@FeedBack_content", SqlDbType.VarChar, 500).Value = FeedBack_content
                        cmdSelectUp.Parameters.Add("@UserID", SqlDbType.Int).Value = CInt(UserID)
                        cmdSelectUp.CommandType = CommandType.StoredProcedure
                        conB.Open()
                        cmdSelectUp.ExecuteNonQuery()
                        conB.Close()




                    End If


                    'Response.Write(Cellular & ":" & smsStatusId & ":" & Len(Cellular) & "<BR>")
                    '   Response.End()
                    'send ....
                End If
            End While
            dr_m.Close()
            cmdSelect.Dispose()
            con.Close()
            Response.Write("<script language='javascript'> {window.opener.location.href = window.opener.location.href; window.close();}</script>")

        End If

    End Sub

End Class